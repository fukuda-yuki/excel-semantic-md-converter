"""Writers for convert output artifacts."""

from __future__ import annotations

import json
import shutil
import tempfile
from collections import defaultdict
from pathlib import Path
from typing import Any

from excel_semantic_md.models import AssetKind, Block, ChartBlock, ImageBlock, ShapeBlock, make_asset_path
from excel_semantic_md.output.models import ConvertOutputFiles, ConvertResult, ConvertSheetResult, PublishedAsset


def write_convert_outputs(result: ConvertResult) -> ConvertOutputFiles:
    staging_dir = Path(
        tempfile.mkdtemp(
            prefix=".excel-semantic-md-staging-",
            dir=str(result.output_dir.parent.resolve()),
        )
    )
    assets_dir = staging_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    try:
        for sheet_result in result.sheets:
            _publish_sheet_assets(staging_dir, result, sheet_result)

        result_markdown = staging_dir / "result.md"
        result_markdown.write_text(_build_result_markdown(result), encoding="utf-8")

        manifest_json = staging_dir / "manifest.json"
        manifest_json.write_text(
            json.dumps(_build_manifest_payload(result), ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

        debug_dir: Path | None = None
        if result.command_options.get("save_debug_json"):
            debug_dir = _write_debug_payloads(staging_dir, result)

        _replace_managed_outputs(staging_dir, result.output_dir)
    except Exception:
        shutil.rmtree(staging_dir, ignore_errors=True)
        raise

    return ConvertOutputFiles(
        result_markdown=result.output_dir / "result.md",
        manifest_json=result.output_dir / "manifest.json",
        assets_dir=result.output_dir / "assets",
        debug_dir=None if not result.command_options.get("save_debug_json") else result.output_dir / "debug",
    )


def _publish_sheet_assets(staging_dir: Path, result: ConvertResult, sheet_result: ConvertSheetResult) -> None:
    sheet_result.assets.clear()
    if sheet_result.render_result is None:
        return

    counters: dict[tuple[str, str], int] = defaultdict(int)
    for artifact in sheet_result.render_result.artifacts:
        if artifact.role == "render_artifact" and not result.command_options.get("save_render_artifacts"):
            continue

        counters[(artifact.block_id, artifact.kind)] += 1
        relative_path = _published_asset_path(
            sheet_index=sheet_result.sheet.sheet_index,
            block_id=artifact.block_id,
            kind=artifact.kind,
            asset_index=counters[(artifact.block_id, artifact.kind)],
            source_path=Path(artifact.path),
        )
        target_path = staging_dir / Path(relative_path)
        target_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(artifact.path, target_path)
        sheet_result.assets.append(
            PublishedAsset(
                sheet_index=sheet_result.sheet.sheet_index,
                sheet_name=sheet_result.sheet.name,
                block_id=artifact.block_id,
                visual_id=artifact.visual_id,
                related_block_id=artifact.related_block_id,
                kind=artifact.kind,
                role=artifact.role,
                source=artifact.source,
                path=relative_path,
                anchor=artifact.anchor,
            )
        )


def _published_asset_path(
    *,
    sheet_index: int,
    block_id: str,
    kind: str,
    asset_index: int,
    source_path: Path,
) -> str:
    asset_kind = _asset_kind_from_value(kind)
    relative = Path(make_asset_path(sheet_index, block_id, asset_kind, asset_index))
    suffix = source_path.suffix or relative.suffix
    return str(relative.with_suffix(suffix).as_posix())


def _asset_kind_from_value(kind: str) -> AssetKind:
    try:
        return AssetKind(kind)
    except ValueError:
        return AssetKind.UNKNOWN


def _build_result_markdown(result: ConvertResult) -> str:
    sections: list[str] = []
    for sheet_result in result.sheets:
        sections.append(f"## {sheet_result.sheet.name}")
        sections.append("")
        if sheet_result.status == "failed":
            sections.extend(_failed_sheet_lines(sheet_result))
        else:
            sections.extend(_successful_sheet_lines(sheet_result))
        sections.append("")
    return "\n".join(sections).strip() + "\n"


def _successful_sheet_lines(sheet_result: ConvertSheetResult) -> list[str]:
    lines: list[str] = []
    markdown = _rewrite_asset_references((sheet_result.markdown or "").strip(), sheet_result.assets)
    if markdown:
        lines.append(markdown)
    else:
        lines.append("_No Markdown content was produced for this sheet._")

    unknown_lines = _unknown_note_lines(sheet_result)
    if unknown_lines:
        if lines:
            lines.append("")
        lines.extend(unknown_lines)

    asset_lines = _markdown_asset_lines(sheet_result, markdown)
    if asset_lines:
        if lines:
            lines.append("")
        lines.extend(asset_lines)
    return lines


def _rewrite_asset_references(markdown: str, assets: list[PublishedAsset]) -> str:
    rewritten = markdown
    for asset in assets:
        basename = Path(asset.path).name
        if basename == asset.path or basename not in rewritten:
            continue
        rewritten = rewritten.replace(basename, asset.path)
    return rewritten


def _failed_sheet_lines(sheet_result: ConvertSheetResult) -> list[str]:
    if not sheet_result.failures:
        return ["Failed to convert this sheet."]
    lines = ["Failed to convert this sheet."]
    for failure in sheet_result.failures:
        lines.append(f"- {_failure_summary(failure)}")
    return lines


def _failure_summary(failure: Any) -> str:
    code = None
    details = getattr(failure, "details", {})
    if isinstance(details, dict):
        code = details.get("code")
    if code:
        return f"{failure.stage}: [{code}] {failure.message}"
    return f"{failure.stage}: {failure.message}"


def _unknown_note_lines(sheet_result: ConvertSheetResult) -> list[str]:
    response = None if sheet_result.llm_result is None else sheet_result.llm_result.response
    if response is None or not response.unknowns:
        return []
    lines: list[str] = []
    for unknown in response.unknowns:
        if isinstance(unknown, str):
            text = unknown
        else:
            text = json.dumps(unknown, ensure_ascii=False, sort_keys=True)
        lines.append(f"> Note: {text}")
    return lines


def _markdown_asset_lines(sheet_result: ConvertSheetResult, markdown: str) -> list[str]:
    referenced_paths = {asset.path for asset in sheet_result.assets if asset.path in markdown}
    lines: list[str] = []
    for asset in sheet_result.assets:
        if asset.role != "markdown":
            continue
        if asset.path in referenced_paths:
            continue
        label = _asset_label(sheet_result.sheet.blocks, asset.block_id)
        lines.append(f"![{label}]({asset.path})")
    return lines


def _asset_label(blocks: list[Block], block_id: str) -> str:
    block = next((item for item in blocks if item.id == block_id), None)
    if isinstance(block, ChartBlock) and block.title:
        return block.title
    if isinstance(block, ImageBlock) and block.alt_text:
        return block.alt_text
    if isinstance(block, ShapeBlock) and block.text:
        return block.text.splitlines()[0][:80]
    return block_id


def _build_manifest_payload(result: ConvertResult) -> dict[str, Any]:
    assets_by_block_id: dict[str, list[PublishedAsset]] = defaultdict(list)
    for sheet_result in result.sheets:
        for asset in sheet_result.assets:
            assets_by_block_id[asset.block_id].append(asset)

    return {
        "schema_version": result.schema_version,
        "input_file_name": result.input_file_name,
        "generated_at": result.generated_at,
        "command_options": dict(result.command_options),
        "sheets": [_sheet_manifest(sheet_result) for sheet_result in result.sheets],
        "blocks": [
            _block_manifest(sheet_result.sheet, block, assets_by_block_id.get(block.id, []))
            for sheet_result in result.sheets
            for block in sheet_result.sheet.blocks
        ],
    }


def _sheet_manifest(sheet_result: ConvertSheetResult) -> dict[str, Any]:
    render_status = "skipped"
    render_warnings: list[dict[str, Any]] = []
    render_failures: list[dict[str, Any]] = []
    if sheet_result.render_result is not None:
        render_status = "failed" if sheet_result.render_result.failures else "succeeded"
        render_warnings = [warning.to_dict() for warning in sheet_result.render_result.warnings]
        render_failures = [failure.to_dict() for failure in sheet_result.render_result.failures]

    llm_payload: dict[str, Any]
    if sheet_result.llm_result is None:
        llm_payload = {"status": "skipped"}
    else:
        llm_payload = {
            "status": sheet_result.llm_result.status,
            "attempts": sheet_result.llm_result.attempts,
        }
        if sheet_result.llm_result.failure is not None:
            llm_payload["failure"] = _failure_payload(sheet_result.llm_result.failure)
        if sheet_result.llm_result.response is not None:
            llm_payload["response"] = {
                "sheet_summary": sheet_result.llm_result.response.sheet_summary,
                "unknowns": list(sheet_result.llm_result.response.unknowns),
            }

    return {
        "sheet_index": sheet_result.sheet.sheet_index,
        "name": sheet_result.sheet.name,
        "status": sheet_result.status,
        "warnings": [_warning_payload(warning) for warning in sheet_result.warnings],
        "failures": [_failure_payload(failure) for failure in sheet_result.failures],
        "render": {
            "status": render_status,
            "warnings": [_warning_payload_dict(warning) for warning in render_warnings],
            "failures": [_failure_payload_dict(failure) for failure in render_failures],
            "assets": [asset.to_dict() for asset in sheet_result.assets],
        },
        "llm": llm_payload,
    }


def _block_manifest(sheet: Any, block: Block, assets: list[PublishedAsset]) -> dict[str, Any]:
    payload = block.to_dict()
    payload["sheet_index"] = sheet.sheet_index
    payload["sheet_name"] = sheet.name
    payload["warnings"] = [_warning_payload_dict(warning) for warning in payload.get("warnings", [])]
    payload["assets"] = [
        {
            "path": asset.path,
            "kind": asset.kind,
            "role": asset.role,
            "source": asset.source,
            "anchor": asset.anchor.to_dict(),
        }
        for asset in assets
    ]
    return payload


def _write_debug_payloads(staging_dir: Path, result: ConvertResult) -> Path:
    debug_dir = staging_dir / "debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    _write_json(debug_dir / "workbook_extraction.json", result.workbook_extraction_payload)
    _write_json(debug_dir / "block_detection.json", result.block_detection_payload)
    _write_json(debug_dir / "linked_blocks.json", result.linked_workbook_payload)
    _write_json(
        debug_dir / "render_plan.json",
        {
            "sheets": [
                sheet_result.render_plan_payload
                for sheet_result in result.sheets
                if sheet_result.render_plan_payload is not None
            ]
        },
    )
    _write_json(
        debug_dir / "llm_input.json",
        {
            "sheets": [
                {
                    "sheet_index": sheet_result.sheet.sheet_index,
                    "name": sheet_result.sheet.name,
                    "input": sheet_result.llm_input_payload,
                }
                for sheet_result in result.sheets
                if sheet_result.llm_input_payload is not None
            ]
        },
    )
    _write_json(
        debug_dir / "llm_response.json",
        {
            "sheets": [
                {
                    "sheet_index": sheet_result.sheet.sheet_index,
                    "name": sheet_result.sheet.name,
                    "result": None if sheet_result.llm_result is None else sheet_result.llm_result.to_dict(),
                }
                for sheet_result in result.sheets
                if sheet_result.llm_result is not None
            ]
        },
    )
    return debug_dir


def _write_json(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _replace_managed_outputs(staging_dir: Path, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for managed_path in (
        output_dir / "assets",
        output_dir / "debug",
        output_dir / "result.md",
        output_dir / "manifest.json",
    ):
        if managed_path.is_dir():
            shutil.rmtree(managed_path)
        elif managed_path.exists():
            managed_path.unlink()

    for item in staging_dir.iterdir():
        shutil.move(str(item), str(output_dir / item.name))
    staging_dir.rmdir()


def _warning_payload(warning: Any) -> dict[str, Any]:
    return {
        "code": warning.code,
        "message": warning.message,
        "details": _sanitize_details(getattr(warning, "details", {})),
    }


def _warning_payload_dict(warning: dict[str, Any]) -> dict[str, Any]:
    return {
        "code": warning["code"],
        "message": warning["message"],
        "details": _sanitize_details(warning.get("details", {})),
    }


def _failure_payload(failure: Any) -> dict[str, Any]:
    return {
        "stage": failure.stage,
        "message": failure.message,
        "details": _sanitize_details(getattr(failure, "details", {})),
    }


def _failure_payload_dict(failure: dict[str, Any]) -> dict[str, Any]:
    return {
        "stage": failure["stage"],
        "message": failure["message"],
        "details": _sanitize_details(failure.get("details", {})),
    }


def _sanitize_details(value: Any, *, key: str | None = None) -> Any:
    if isinstance(value, dict):
        return {subkey: _sanitize_details(subvalue, key=subkey) for subkey, subvalue in value.items()}
    if isinstance(value, list):
        return [_sanitize_details(item, key=key) for item in value]
    if isinstance(value, str):
        if key in {"path", "workbook", "temp_dir"} and _looks_local_path(value):
            return "[redacted]"
        if "excel-semantic-md-render-" in value or "excel-semantic-md-staging-" in value:
            return "[redacted]"
    return value


def _looks_local_path(value: str) -> bool:
    if value.startswith(("/", "\\")):
        return True
    return len(value) > 2 and value[1] == ":" and value[2] in {"\\", "/"}
