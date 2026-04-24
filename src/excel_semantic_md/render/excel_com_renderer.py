"""Excel COM rendering helpers for live confirmation."""

from __future__ import annotations

import contextlib
import platform
import tempfile
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Any

from excel_semantic_md.models import FailureInfo, Rect, WarningInfo
from excel_semantic_md.render.types import RenderArtifact, RenderPlanItem, RenderSheetResult


@dataclass(frozen=True)
class _ComCandidate:
    object_ref: Any
    rect: Rect
    hint_text: str | None
    hint_alt_text: str | None


class RenderTaskError(RuntimeError):
    def __init__(self, message: str, *, details: dict[str, Any]) -> None:
        super().__init__(message)
        self.details = details


class ExcelSession:
    """Own a dedicated Excel COM session for one rendering job."""

    def __init__(self, workbook_path: Path) -> None:
        self.workbook_path = workbook_path
        self.cleanup_warnings: list[WarningInfo] = []
        self._pythoncom: Any | None = None
        self._win32_client: Any | None = None
        self._com_initialized = False
        self._cleaned_up = False
        self._previous_automation_security: Any | None = None
        self.app: Any | None = None
        self.workbook: Any | None = None

    def __enter__(self) -> "ExcelSession":
        self._pythoncom, self._win32_client = _load_excel_com_modules()
        self._pythoncom.CoInitialize()
        self._com_initialized = True
        try:
            self.app = self._win32_client.DispatchEx("Excel.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self._previous_automation_security = getattr(self.app, "AutomationSecurity", None)
            try:
                self.app.AutomationSecurity = 3
            except Exception as exc:
                if self.workbook_path.suffix.lower() == ".xlsm":
                    raise RuntimeError(
                        "Failed to force Excel macro-disabled automation security for .xlsm workbook."
                    ) from exc
            self.workbook = self.app.Workbooks.Open(
                str(self.workbook_path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                Notify=False,
            )
        except Exception:
            self._cleanup_session()
            raise
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        self._cleanup_session()

    def _cleanup_session(self) -> None:
        if self._cleaned_up:
            return
        self._cleaned_up = True
        if self.app is not None and self._previous_automation_security is not None:
            try:
                self.app.AutomationSecurity = self._previous_automation_security
            except Exception:
                pass
        if self.workbook is not None:
            try:
                self.workbook.Close(False)
            except Exception as close_exc:
                self.cleanup_warnings.append(
                    WarningInfo(
                        code="excel_workbook_close_failed",
                        message="Excel workbook cleanup failed after rendering.",
                        details={"error": str(close_exc), "workbook": self.workbook_path.name},
                    )
                )
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception as quit_exc:
                self.cleanup_warnings.append(
                    WarningInfo(
                        code="excel_application_quit_failed",
                        message="Excel application cleanup failed after rendering.",
                        details={"error": str(quit_exc), "workbook": self.workbook_path.name},
                    )
                )
        if self._pythoncom is not None and self._com_initialized:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_initialized = False

    def worksheet(self, sheet_name: str) -> Any:
        if self.workbook is None:
            raise RuntimeError("Excel session is not open.")
        return self.workbook.Worksheets(sheet_name)


def render_with_excel_com(
    workbook_path: Path,
    *,
    input_file_name: str,
    sheet_name: str,
    plan_items: list[RenderPlanItem],
    warnings: list[WarningInfo],
    failures: list[FailureInfo],
) -> RenderSheetResult:
    """Render one sheet using Excel COM and return JSON-ready results."""

    temp_dir = Path(tempfile.mkdtemp(prefix="excel-semantic-md-render-")).resolve()
    result = RenderSheetResult(
        input_file_name=input_file_name,
        sheet_name=sheet_name,
        temp_dir=str(temp_dir),
        warnings=list(warnings),
        failures=list(failures),
    )

    available, message = excel_com_diagnostic()
    if not available:
        result.failures.append(
            FailureInfo(
                stage="render",
                message=message,
                details={"sheet": sheet_name},
            )
        )
        return result

    session: ExcelSession | None = None
    try:
        session = ExcelSession(workbook_path)
        with session:
            worksheet = session.worksheet(sheet_name)
            output_dir = temp_dir / f"sheet-{_safe_component(sheet_name)}"
            output_dir.mkdir(parents=True, exist_ok=True)
            counters: dict[tuple[str, str, str], int] = defaultdict(int)
            for item in plan_items:
                try:
                    artifact = _render_plan_item(
                        workbook_path=workbook_path,
                        worksheet=worksheet,
                        item=item,
                        output_dir=output_dir,
                        counters=counters,
                    )
                except RenderTaskError as exc:
                    result.failures.append(
                        FailureInfo(
                            stage="render",
                            message=str(exc),
                            details=exc.details,
                        )
                    )
                    continue
                except Exception as exc:
                    result.failures.append(
                        FailureInfo(
                            stage="render",
                            message="Render artifact failed unexpectedly.",
                            details={
                                "block_id": item.block.id,
                                "kind": item.kind,
                                "error": str(exc),
                            },
                        )
                    )
                    continue
                result.artifacts.append(artifact)
    except Exception as exc:
        result.failures.append(
            FailureInfo(
                stage="render",
                message="Excel COM rendering failed unexpectedly.",
                details={"sheet": sheet_name, "error": str(exc)},
            )
        )
    finally:
        if session is not None:
            result.warnings.extend(session.cleanup_warnings)

    return result


def excel_com_diagnostic() -> tuple[bool, str]:
    if platform.system() != "Windows":
        return False, "Excel COM rendering requires Windows."
    try:
        _load_excel_com_modules()
    except ImportError:
        return False, "Excel COM rendering requires pywin32 modules (`pythoncom` and `win32com.client`)."
    return True, "Excel COM is available."


def _load_excel_com_modules() -> tuple[Any, Any]:
    import pythoncom
    import win32com.client

    return pythoncom, win32com.client


def _render_plan_item(
    *,
    workbook_path: Path,
    worksheet: Any,
    item: RenderPlanItem,
    output_dir: Path,
    counters: dict[tuple[str, str, str], int],
) -> RenderArtifact:
    output_path = _next_output_path(output_dir, item, counters)
    try:
        if item.source == "ooxml_image_copy":
            _copy_package_part(workbook_path, item.target_part, output_path)
        elif item.kind == "range":
            range_ref = worksheet.Range(item.block.anchor.a1)
            _copy_object_to_png(worksheet, range_ref, output_path)
        elif item.kind == "chart":
            chart_object = _match_chart_object(worksheet, item)
            exported = chart_object.Chart.Export(str(output_path), "PNG")
            if exported is False:
                raise RenderTaskError(
                    "Chart.Export returned False.",
                    details={"block_id": item.block.id, "path": output_path.name},
                )
        elif item.kind in {"shape", "image"}:
            shape = _match_shape_object(worksheet, item)
            _copy_object_to_png(worksheet, shape, output_path)
        else:
            raise RenderTaskError(
                "Render plan item kind is not supported.",
                details={"block_id": item.block.id, "kind": item.kind},
            )
    except Exception:
        with contextlib.suppress(OSError):
            output_path.unlink(missing_ok=True)
        raise

    return RenderArtifact(
        block_id=item.block.id,
        visual_id=item.block.visual_id,
        related_block_id=item.block.related_block_id,
        kind=item.kind,
        role=item.role.value,
        path=str(output_path),
        source=item.source,
        anchor=item.block.anchor,
    )


def _next_output_path(
    output_dir: Path,
    item: RenderPlanItem,
    counters: dict[tuple[str, str, str], int],
) -> Path:
    key = (item.block.id, item.kind, item.role.value)
    counters[key] += 1
    if item.source == "ooxml_image_copy":
        suffix = PurePosixPath(item.target_part or "").suffix or ".bin"
    else:
        suffix = ".png"
    return output_dir / f"{item.block.id}-{item.kind}-{item.role.value}-{counters[key]:03d}{suffix}"


def _copy_package_part(workbook_path: Path, target_part: str | None, output_path: Path) -> None:
    if target_part is None:
        raise RenderTaskError(
            "Original image asset copy is missing an OOXML target part.",
            details={"path": output_path.name},
        )
    with zipfile.ZipFile(workbook_path) as archive:
        try:
            payload = archive.read(target_part)
        except KeyError as exc:
            raise RenderTaskError(
                "Original image asset part was not found in the workbook package.",
                details={"target_part": target_part, "path": output_path.name},
            ) from exc
    output_path.write_bytes(payload)


def _copy_object_to_png(worksheet: Any, excel_object: Any, output_path: Path) -> None:
    temp_chart = worksheet.ChartObjects().Add(0, 0, _object_width(excel_object), _object_height(excel_object))
    try:
        excel_object.CopyPicture()
        temp_chart.Chart.Paste()
        exported = temp_chart.Chart.Export(str(output_path), "PNG")
        if exported is False:
            raise RenderTaskError(
                "Clipboard picture export returned False.",
                details={"path": output_path.name},
            )
    finally:
        try:
            temp_chart.Delete()
        except Exception:
            pass


def _match_chart_object(worksheet: Any, item: RenderPlanItem) -> Any:
    candidates = _chart_candidates(worksheet, item.block.anchor.sheet)
    return _choose_candidate(
        candidates,
        expected_rect=item.block.anchor,
        block_id=item.block.id,
        kind="chart",
        text_hint=getattr(item.block, "title", None),
        alt_hint=None,
    ).object_ref


def _match_shape_object(worksheet: Any, item: RenderPlanItem) -> Any:
    candidates = _shape_candidates(worksheet, item.block.anchor.sheet)
    return _choose_candidate(
        candidates,
        expected_rect=item.block.anchor,
        block_id=item.block.id,
        kind=item.kind,
        text_hint=getattr(item.block, "text", None),
        alt_hint=getattr(item.block, "alt_text", None),
    ).object_ref


def _chart_candidates(worksheet: Any, sheet_name: str) -> list[_ComCandidate]:
    candidates: list[_ComCandidate] = []
    for chart_object in _iter_collection(worksheet.ChartObjects()):
        rect = _object_rect(chart_object, sheet_name)
        if rect is None:
            continue
        title = _chart_title(chart_object)
        candidates.append(
            _ComCandidate(
                object_ref=chart_object,
                rect=rect,
                hint_text=title,
                hint_alt_text=None,
            )
        )
    return candidates


def _shape_candidates(worksheet: Any, sheet_name: str) -> list[_ComCandidate]:
    candidates: list[_ComCandidate] = []
    for shape in _iter_collection(worksheet.Shapes):
        rect = _object_rect(shape, sheet_name)
        if rect is None:
            continue
        candidates.append(
            _ComCandidate(
                object_ref=shape,
                rect=rect,
                hint_text=_shape_text(shape),
                hint_alt_text=getattr(shape, "AlternativeText", None),
            )
        )
    return candidates


def _choose_candidate(
    candidates: list[_ComCandidate],
    *,
    expected_rect: Rect,
    block_id: str,
    kind: str,
    text_hint: str | None,
    alt_hint: str | None,
) -> _ComCandidate:
    if not candidates:
        raise RenderTaskError(
            "No Excel COM object matched the renderable block.",
            details={"block_id": block_id, "kind": kind, "anchor": expected_rect.a1},
        )

    exact = [candidate for candidate in candidates if _rect_sort_key(candidate.rect) == _rect_sort_key(expected_rect)]
    pool = exact if exact else candidates
    if not exact:
        min_distance = min(_rect_distance(expected_rect, candidate.rect) for candidate in pool)
        pool = [candidate for candidate in pool if _rect_distance(expected_rect, candidate.rect) == min_distance]

    hinted = _apply_hints(pool, text_hint=text_hint, alt_hint=alt_hint)
    if len(hinted) == 1:
        return hinted[0]
    if len(hinted) > 1:
        pool = hinted

    if len(pool) == 1:
        return pool[0]

    raise RenderTaskError(
        "Multiple Excel COM objects matched the renderable block.",
        details={
            "block_id": block_id,
            "kind": kind,
            "anchor": expected_rect.a1,
            "candidate_count": len(pool),
        },
    )


def _apply_hints(
    candidates: list[_ComCandidate],
    *,
    text_hint: str | None,
    alt_hint: str | None,
) -> list[_ComCandidate]:
    if len(candidates) <= 1:
        return candidates

    filtered = candidates
    normalized_text = _normalize_hint(text_hint)
    normalized_alt = _normalize_hint(alt_hint)
    if normalized_text is not None:
        text_matches = [
            candidate
            for candidate in filtered
            if _hint_matches(candidate.hint_text, normalized_text)
        ]
        if text_matches:
            filtered = text_matches
    if len(filtered) <= 1:
        return filtered
    if normalized_alt is not None:
        alt_matches = [
            candidate
            for candidate in filtered
            if _hint_matches(candidate.hint_alt_text, normalized_alt)
        ]
        if alt_matches:
            filtered = alt_matches
    return filtered


def _hint_matches(candidate_value: str | None, expected_normalized: str) -> bool:
    candidate_normalized = _normalize_hint(candidate_value)
    if candidate_normalized is None:
        return False
    return candidate_normalized == expected_normalized or expected_normalized in candidate_normalized


def _normalize_hint(value: str | None) -> str | None:
    if value is None:
        return None
    stripped = " ".join(value.split()).strip().lower()
    return stripped or None


def _iter_collection(collection: Any) -> list[Any]:
    if isinstance(collection, list):
        return collection
    count = int(getattr(collection, "Count", 0))
    return [collection.Item(index) for index in range(1, count + 1)]


def _object_rect(com_object: Any, sheet_name: str) -> Rect | None:
    try:
        top_left = com_object.TopLeftCell
        bottom_right = com_object.BottomRightCell
    except Exception:
        return None
    try:
        start_row = int(top_left.Row)
        start_col = int(top_left.Column)
        end_row = int(bottom_right.Row)
        end_col = int(bottom_right.Column)
    except Exception:
        return None
    return Rect(
        sheet=sheet_name,
        start_row=min(start_row, end_row),
        start_col=min(start_col, end_col),
        end_row=max(start_row, end_row),
        end_col=max(start_col, end_col),
        a1=_rect_a1(start_row=min(start_row, end_row), start_col=min(start_col, end_col), end_row=max(start_row, end_row), end_col=max(start_col, end_col)),
    )


def _chart_title(chart_object: Any) -> str | None:
    try:
        chart = chart_object.Chart
        if getattr(chart, "HasTitle", False):
            return str(chart.ChartTitle.Text)
    except Exception:
        return None
    return None


def _shape_text(shape: Any) -> str | None:
    for path in (
        ("TextFrame2", "TextRange", "Text"),
        ("TextFrame", "Characters"),
    ):
        try:
            value = shape
            for attr in path:
                value = getattr(value, attr)
                if callable(value):
                    value = value()
            text = str(value)
            if text.strip():
                return text
        except Exception:
            continue
    return None


def _object_width(excel_object: Any) -> float:
    try:
        width = float(excel_object.Width)
    except Exception:
        width = 240.0
    return max(width, 1.0)


def _object_height(excel_object: Any) -> float:
    try:
        height = float(excel_object.Height)
    except Exception:
        height = 120.0
    return max(height, 1.0)


def _rect_sort_key(anchor: Rect) -> tuple[int, int, int, int]:
    return (anchor.start_row, anchor.start_col, anchor.end_row, anchor.end_col)


def _rect_distance(left: Rect, right: Rect) -> int:
    row_gap = max(0, left.start_row - right.end_row - 1, right.start_row - left.end_row - 1)
    col_gap = max(0, left.start_col - right.end_col - 1, right.start_col - left.end_col - 1)
    return max(row_gap, col_gap)


def _rect_a1(*, start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    start = f"{_column_letters(start_col)}{start_row}"
    end = f"{_column_letters(end_col)}{end_row}"
    return start if start == end else f"{start}:{end}"


def _column_letters(col: int) -> str:
    letters: list[str] = []
    remaining = col
    while remaining > 0:
        remaining, offset = divmod(remaining - 1, 26)
        letters.append(chr(ord("A") + offset))
    return "".join(reversed(letters))


def _safe_component(value: str) -> str:
    return "".join(char if char.isalnum() or char in {"-", "_"} else "_" for char in value) or "sheet"
