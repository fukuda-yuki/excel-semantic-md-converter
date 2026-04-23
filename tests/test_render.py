from __future__ import annotations

import contextlib
import io
import json
import shutil
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import excel_semantic_md.cli.main as cli_main
import pytest
from excel_semantic_md.excel import detect_blocks, link_visuals, read_visual_metadata, read_workbook
from excel_semantic_md.models import AssetRole, ChartBlock, ParagraphBlock, Rect, SourceKind
from excel_semantic_md.render.excel_com_renderer import ExcelSession, render_with_excel_com
from excel_semantic_md.render.planner import build_render_plan
from excel_semantic_md.render.types import RenderArtifact, RenderPlanItem, RenderSheetResult


FIXTURES = Path(__file__).resolve().parent / "fixtures" / "visuals"


def _run_cli(argv: list[str]) -> tuple[int | str | None, str, str]:
    stdout = io.StringIO()
    stderr = io.StringIO()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        try:
            code = cli_main.main(argv)
        except SystemExit as exc:
            code = exc.code
    return code, stdout.getvalue(), stderr.getvalue()


def test_build_render_plan_for_image_sheet_includes_primary_and_original_asset() -> None:
    workbook = _linked_workbook("image-visual.xlsx")
    sheet = workbook.sheets[0]
    visual_sheet = read_visual_metadata(FIXTURES / "image-visual.xlsx")[0]

    items, warnings, failures = build_render_plan(sheet, visual_sheet)

    assert warnings == []
    assert failures == []
    assert [(item.block.id, item.kind, item.role.value, item.source) for item in items] == [
        ("s001-b001-paragraph", "range", "render_artifact", "range_copy_picture"),
        ("s001-b002-image", "image", "markdown", "ooxml_image_copy"),
        ("s001-b002-image", "image", "render_artifact", "shape_copy_picture"),
    ]


def test_build_render_plan_adds_auxiliary_range_artifact_when_enabled() -> None:
    workbook = _linked_workbook("table-shape-visual.xlsx")
    sheet = workbook.sheets[0]
    visual_sheet = read_visual_metadata(FIXTURES / "table-shape-visual.xlsx")[0]

    items, warnings, failures = build_render_plan(sheet, visual_sheet, save_render_artifacts=True)

    assert warnings == []
    assert failures == []
    assert [(item.kind, item.role.value, item.source) for item in items] == [
        ("range", "render_artifact", "range_copy_picture"),
        ("shape", "markdown", "shape_copy_picture"),
        ("range", "render_artifact", "range_copy_picture"),
    ]


def test_render_with_excel_com_returns_handled_failure_when_com_is_unavailable() -> None:
    plan_item = RenderPlanItem(
        block=ParagraphBlock(
            id="s001-b001-paragraph",
            anchor=Rect(sheet="Plain", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
            source=SourceKind.CELLS,
            text="Plain text",
        ),
        kind="range",
        role=AssetRole.MARKDOWN,
        source="range_copy_picture",
    )

    with mock.patch("excel_semantic_md.render.excel_com_renderer.excel_com_diagnostic", return_value=(False, "Excel COM unavailable for test.")):
        result = render_with_excel_com(
            FIXTURES / "no-visuals.xlsx",
            input_file_name="no-visuals.xlsx",
            sheet_name="Plain",
            plan_items=[plan_item],
            warnings=[],
            failures=[],
        )

    try:
        assert result.artifacts == []
        assert result.failures[0].stage == "render"
        assert result.failures[0].message == "Excel COM unavailable for test."
        assert Path(result.temp_dir).is_dir()
    finally:
        shutil.rmtree(result.temp_dir, ignore_errors=True)


def test_excel_session_cleans_up_when_workbook_open_fails() -> None:
    pythoncom = SimpleNamespace(CoInitialize=mock.Mock(), CoUninitialize=mock.Mock())
    app = SimpleNamespace(
        Visible=True,
        DisplayAlerts=True,
        AutomationSecurity=1,
        Workbooks=SimpleNamespace(Open=mock.Mock(side_effect=RuntimeError("open failed"))),
        Quit=mock.Mock(),
    )
    win32_client = SimpleNamespace(DispatchEx=mock.Mock(return_value=app))

    session = ExcelSession(FIXTURES / "image-visual.xlsm")
    with mock.patch("excel_semantic_md.render.excel_com_renderer._load_excel_com_modules", return_value=(pythoncom, win32_client)):
        try:
            session.__enter__()
        except RuntimeError as exc:
            assert str(exc) == "open failed"
        else:
            raise AssertionError("ExcelSession.__enter__() should have raised.")

    assert win32_client.DispatchEx.call_count == 1
    assert app.AutomationSecurity == 1
    assert app.Quit.call_count == 1
    assert pythoncom.CoInitialize.call_count == 1
    assert pythoncom.CoUninitialize.call_count == 1


def test_excel_session_fails_closed_for_xlsm_when_macro_security_cannot_be_set() -> None:
    pythoncom = SimpleNamespace(CoInitialize=mock.Mock(), CoUninitialize=mock.Mock())
    app = _automation_security_failing_app()
    win32_client = SimpleNamespace(DispatchEx=mock.Mock(return_value=app))

    session = ExcelSession(FIXTURES / "image-visual.xlsm")
    with mock.patch("excel_semantic_md.render.excel_com_renderer._load_excel_com_modules", return_value=(pythoncom, win32_client)):
        with pytest.raises(RuntimeError, match="macro-disabled"):
            session.__enter__()

    assert app.Workbooks.Open.call_count == 0
    assert app.Quit.call_count == 1
    assert pythoncom.CoInitialize.call_count == 1
    assert pythoncom.CoUninitialize.call_count == 1


def test_excel_session_sets_macro_disabled_and_reports_cleanup_failures() -> None:
    pythoncom = SimpleNamespace(CoInitialize=mock.Mock(), CoUninitialize=mock.Mock())
    workbook = SimpleNamespace(Close=mock.Mock(side_effect=RuntimeError("close failed")))
    app = SimpleNamespace(
        Visible=True,
        DisplayAlerts=True,
        AutomationSecurity=1,
        Workbooks=SimpleNamespace(Open=mock.Mock(return_value=workbook)),
        Quit=mock.Mock(side_effect=RuntimeError("quit failed")),
    )
    win32_client = SimpleNamespace(DispatchEx=mock.Mock(return_value=app))

    session = ExcelSession(FIXTURES / "image-visual.xlsm")
    with mock.patch("excel_semantic_md.render.excel_com_renderer._load_excel_com_modules", return_value=(pythoncom, win32_client)):
        entered = session.__enter__()
        assert entered is session
        assert app.Visible is False
        assert app.DisplayAlerts is False
        assert workbook is session.workbook
        session.__exit__(None, None, None)

    assert pythoncom.CoInitialize.call_count == 1
    assert pythoncom.CoUninitialize.call_count == 1
    warning_codes = [warning.code for warning in session.cleanup_warnings]
    assert warning_codes == ["excel_workbook_close_failed", "excel_application_quit_failed"]
    assert [warning.details["workbook"] for warning in session.cleanup_warnings] == ["image-visual.xlsm", "image-visual.xlsm"]


def test_render_with_excel_com_executes_chart_export_and_original_image_copy() -> None:
    chart_block = ChartBlock(
        id="s001-b001-chart",
        anchor=Rect(sheet="Chart", start_row=2, start_col=4, end_row=2, end_col=4, a1="D2"),
        source=SourceKind.CHART,
        visual_id="s001-v001-chart",
        title="Quarterly Sales",
    )
    plan_items = [
        RenderPlanItem(block=chart_block, kind="chart", role=AssetRole.MARKDOWN, source="chart_export"),
        RenderPlanItem(
            block=ChartBlock(
                id="s001-b003-image",
                anchor=Rect(sheet="Chart", start_row=3, start_col=5, end_row=4, end_col=6, a1="E3:F4"),
                source=SourceKind.IMAGE,
                visual_id="s001-v002-image",
            ),
            kind="image",
            role=AssetRole.RENDER_ARTIFACT,
            source="ooxml_image_copy",
            target_part="xl/media/image1.png",
        ),
    ]

    chart_object = SimpleNamespace(
        TopLeftCell=SimpleNamespace(Row=2, Column=4),
        BottomRightCell=SimpleNamespace(Row=2, Column=4),
        Chart=SimpleNamespace(
            HasTitle=True,
            ChartTitle=SimpleNamespace(Text="Quarterly Sales"),
            Export=mock.Mock(return_value=True),
        ),
    )
    worksheet = SimpleNamespace(
        ChartObjects=mock.Mock(return_value=[chart_object]),
        Shapes=[],
    )
    session = mock.MagicMock()
    session.__enter__.return_value = session
    session.__exit__.return_value = None
    session.worksheet.return_value = worksheet
    session.cleanup_warnings = []

    with (
        mock.patch("excel_semantic_md.render.excel_com_renderer.excel_com_diagnostic", return_value=(True, "ok")),
        mock.patch("excel_semantic_md.render.excel_com_renderer.ExcelSession", return_value=session),
    ):
        result = render_with_excel_com(
            FIXTURES / "image-visual.xlsx",
            input_file_name="image-visual.xlsx",
            sheet_name="Chart",
            plan_items=plan_items,
            warnings=[],
            failures=[],
        )

    try:
        assert result.failures == []
        assert [artifact.source for artifact in result.artifacts] == ["chart_export", "ooxml_image_copy"]
        chart_object.Chart.Export.assert_called_once()
        copied_image = Path(result.artifacts[1].path)
        assert copied_image.is_file()
        assert copied_image.read_bytes()
    finally:
        shutil.rmtree(result.temp_dir, ignore_errors=True)


def test_render_with_excel_com_failure_details_do_not_include_absolute_artifact_path() -> None:
    chart_block = ChartBlock(
        id="s001-b001-chart",
        anchor=Rect(sheet="Chart", start_row=2, start_col=4, end_row=2, end_col=4, a1="D2"),
        source=SourceKind.CHART,
        visual_id="s001-v001-chart",
        title="Quarterly Sales",
    )
    plan_item = RenderPlanItem(block=chart_block, kind="chart", role=AssetRole.MARKDOWN, source="chart_export")
    chart_object = SimpleNamespace(
        TopLeftCell=SimpleNamespace(Row=2, Column=4),
        BottomRightCell=SimpleNamespace(Row=2, Column=4),
        Chart=SimpleNamespace(
            HasTitle=True,
            ChartTitle=SimpleNamespace(Text="Quarterly Sales"),
            Export=mock.Mock(return_value=False),
        ),
    )
    worksheet = SimpleNamespace(
        ChartObjects=mock.Mock(return_value=[chart_object]),
        Shapes=[],
    )
    session = mock.MagicMock()
    session.__enter__.return_value = session
    session.__exit__.return_value = None
    session.worksheet.return_value = worksheet
    session.cleanup_warnings = []

    with (
        mock.patch("excel_semantic_md.render.excel_com_renderer.excel_com_diagnostic", return_value=(True, "ok")),
        mock.patch("excel_semantic_md.render.excel_com_renderer.ExcelSession", return_value=session),
    ):
        result = render_with_excel_com(
            FIXTURES / "chart-visual.xlsx",
            input_file_name="chart-visual.xlsx",
            sheet_name="Chart",
            plan_items=[plan_item],
            warnings=[],
            failures=[],
        )

    try:
        failure_path = result.failures[0].details["path"]
        assert failure_path.endswith(".png")
        assert "\\" not in failure_path
        assert "/" not in failure_path
    finally:
        shutil.rmtree(result.temp_dir, ignore_errors=True)


def test_render_with_excel_com_records_ambiguous_shape_match_as_failure() -> None:
    shape_one = _shape_candidate("Quarterly callout", "Logo", 2, 2, 4, 4)
    shape_two = _shape_candidate("Quarterly callout", "Logo", 2, 2, 4, 4)
    worksheet = SimpleNamespace(
        ChartObjects=mock.Mock(return_value=SimpleNamespace(Add=mock.Mock())),
        Shapes=[shape_one, shape_two],
    )
    session = mock.MagicMock()
    session.__enter__.return_value = session
    session.__exit__.return_value = None
    session.worksheet.return_value = worksheet
    session.cleanup_warnings = []
    block = SimpleNamespace(
        id="s001-b001-shape",
        visual_id="s001-v001-shape",
        related_block_id="s001-b000-paragraph",
        anchor=Rect(sheet="Sheet1", start_row=2, start_col=2, end_row=4, end_col=4, a1="B2:D4"),
        text="Quarterly callout",
        alt_text=None,
    )
    plan_item = RenderPlanItem(block=block, kind="shape", role=AssetRole.MARKDOWN, source="shape_copy_picture")

    with (
        mock.patch("excel_semantic_md.render.excel_com_renderer.excel_com_diagnostic", return_value=(True, "ok")),
        mock.patch("excel_semantic_md.render.excel_com_renderer.ExcelSession", return_value=session),
    ):
        result = render_with_excel_com(
            FIXTURES / "shape-unknown.xlsx",
            input_file_name="shape-unknown.xlsx",
            sheet_name="Sheet1",
            plan_items=[plan_item],
            warnings=[],
            failures=[],
        )

    try:
        assert result.artifacts == []
        assert result.failures[0].details["candidate_count"] == 2
    finally:
        shutil.rmtree(result.temp_dir, ignore_errors=True)


def test_render_command_outputs_json_when_rendering_cannot_run(monkeypatch) -> None:
    code, stdout, stderr = _run_cli(["render", "--input", str(FIXTURES / "no-visuals.xlsx"), "--sheet", "Plain"])

    payload = json.loads(stdout)
    assert code == 1
    assert stderr == ""
    assert payload["input_file_name"] == "no-visuals.xlsx"
    assert payload["sheet_name"] == "Plain"
    assert payload["failures"]
    shutil.rmtree(payload["temp_dir"], ignore_errors=True)


def test_render_command_success_outputs_contract_without_convert_outputs(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    artifact_path = (tmp_path / "rendered-range.png").resolve()

    def fake_render_with_excel_com(
        _input_path: Path,
        *,
        input_file_name: str,
        sheet_name: str,
        plan_items: list[RenderPlanItem],
        warnings,
        failures,
    ) -> RenderSheetResult:
        temp_dir = artifact_path.parent
        artifact_path.write_bytes(b"png")
        return RenderSheetResult(
            input_file_name=input_file_name,
            sheet_name=sheet_name,
            temp_dir=str(temp_dir),
            artifacts=[
                RenderArtifact(
                    block_id=plan_items[0].block.id,
                    visual_id=None,
                    related_block_id=None,
                    kind="range",
                    role="render_artifact",
                    path=str(artifact_path),
                    source="range_copy_picture",
                    anchor=plan_items[0].block.anchor,
                )
            ],
            warnings=list(warnings),
            failures=list(failures),
        )

    monkeypatch.setattr("excel_semantic_md.render.render_with_excel_com", fake_render_with_excel_com)
    monkeypatch.setattr(
        "excel_semantic_md.llm.adapter.GitHubCopilotSdkAdapter.run_sheet",
        mock.Mock(side_effect=AssertionError("render must not invoke LLM")),
    )

    code, stdout, stderr = _run_cli(["render", "--input", str(FIXTURES / "no-visuals.xlsx"), "--sheet", "Plain"])
    payload = json.loads(stdout)

    try:
        assert code == 0
        assert stderr == ""
        assert set(payload) >= {"input_file_name", "sheet_name", "temp_dir", "artifacts", "warnings", "failures"}
        assert payload["artifacts"] == [
            {
                "block_id": "s001-b001-paragraph",
                "visual_id": None,
                "related_block_id": None,
                "kind": "range",
                "role": "render_artifact",
                "path": str(artifact_path),
                "source": "range_copy_picture",
                "anchor": {
                    "sheet": "Plain",
                    "start_row": 1,
                    "start_col": 1,
                    "end_row": 1,
                    "end_col": 1,
                    "a1": "A1",
                },
            }
        ]
        assert payload["failures"] == []
        assert not (Path(payload["temp_dir"]) / "result.md").exists()
        assert not (Path(payload["temp_dir"]) / "manifest.json").exists()
    finally:
        artifact_path.unlink(missing_ok=True)


def test_render_command_rejects_unknown_sheet_name() -> None:
    code, _stdout, stderr = _run_cli(["render", "--input", str(FIXTURES / "no-visuals.xlsx"), "--sheet", "Missing"])

    assert code == 2
    assert "sheet was not found among visible workbook sheets" in stderr


def _linked_workbook(name: str):
    path = FIXTURES / name
    return link_visuals(detect_blocks(read_workbook(path)), read_visual_metadata(path))


def _shape_candidate(text: str, alt_text: str, start_row: int, start_col: int, end_row: int, end_col: int):
    return SimpleNamespace(
        TopLeftCell=SimpleNamespace(Row=start_row, Column=start_col),
        BottomRightCell=SimpleNamespace(Row=end_row, Column=end_col),
        AlternativeText=alt_text,
        TextFrame2=SimpleNamespace(TextRange=SimpleNamespace(Text=text)),
        CopyPicture=mock.Mock(),
        Width=120,
        Height=80,
    )


def _automation_security_failing_app():
    class App:
        def __init__(self) -> None:
            object.__setattr__(self, "Visible", True)
            object.__setattr__(self, "DisplayAlerts", True)
            object.__setattr__(self, "AutomationSecurity", 1)
            object.__setattr__(self, "Workbooks", SimpleNamespace(Open=mock.Mock()))
            object.__setattr__(self, "Quit", mock.Mock())
            object.__setattr__(self, "_fail_automation_security", True)

        def __setattr__(self, name: str, value) -> None:
            if name == "AutomationSecurity" and getattr(self, "_fail_automation_security", False):
                raise RuntimeError("automation security failed")
            object.__setattr__(self, name, value)

    return App()
