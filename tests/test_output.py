from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

import excel_semantic_md.cli.main as cli_main
import excel_semantic_md.output.writers as output_writers
import pytest

from excel_semantic_md.app.convert_pipeline import run_convert_pipeline
from excel_semantic_md.llm.models import LlmResponse, LlmRunResult
from excel_semantic_md.models import (
    ChartBlock,
    FailureInfo,
    ParagraphBlock,
    Rect,
    SheetModel,
    SourceKind,
    WarningInfo,
)
from excel_semantic_md.output import ConvertResult, ConvertSheetResult, write_convert_outputs
from excel_semantic_md.render.types import RenderArtifact, RenderPlanItem, RenderSheetResult


def test_write_convert_outputs_writes_result_and_manifest_with_failed_sheet(tmp_path: Path) -> None:
    success_sheet = SheetModel(
        sheet_index=1,
        name="Summary",
        blocks=[
            ParagraphBlock(
                id="s001-b001-paragraph",
                anchor=Rect(sheet="Summary", start_row=1, start_col=1, end_row=1, end_col=2, a1="A1:B1"),
                source=SourceKind.CELLS,
                text="Intro",
            ),
            ChartBlock(
                id="s001-b002-chart",
                anchor=Rect(sheet="Summary", start_row=3, start_col=1, end_row=8, end_col=6, a1="A3:F8"),
                source=SourceKind.CHART,
                visual_id="s001-v001-chart",
                related_block_id="s001-b001-paragraph",
                title="Revenue chart",
            ),
        ],
    )
    failed_sheet = SheetModel(
        sheet_index=2,
        name="Broken",
        blocks=[],
        failures=[FailureInfo(stage="workbook_reading", message="Formula cache missing.", details={"code": "formula_cached_value_missing"})],
    )
    render_result = RenderSheetResult(
        input_file_name="book.xlsx",
        sheet_name="Summary",
        temp_dir=str(tmp_path / "render-summary"),
        artifacts=[
            _artifact(
                tmp_path / "render-summary" / "chart.png",
                block_id="s001-b002-chart",
                kind="chart",
                role="markdown",
                source="chart_export",
                anchor=Rect(sheet="Summary", start_row=3, start_col=1, end_row=8, end_col=6, a1="A3:F8"),
                related_block_id="s001-b001-paragraph",
                visual_id="s001-v001-chart",
            ),
            _artifact(
                tmp_path / "render-summary" / "range.png",
                block_id="s001-b001-paragraph",
                kind="range",
                role="render_artifact",
                source="range_copy_picture",
                anchor=Rect(sheet="Summary", start_row=1, start_col=1, end_row=1, end_col=2, a1="A1:B1"),
                related_block_id=None,
                visual_id=None,
            ),
        ],
    )
    convert_result = ConvertResult(
        input_file_name="book.xlsx",
        schema_version="phase1.0",
        generated_at="2026-04-23T00:00:00+00:00",
        command_options={
            "model": "text-model",
            "vision_model": "vision-model",
            "max_images_per_sheet": 2,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
        output_dir=tmp_path / "out",
        workbook_extraction_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        block_detection_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        linked_workbook_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        sheets=[
            ConvertSheetResult(
                sheet=success_sheet,
                status="succeeded",
                warnings=[WarningInfo(code="sheet_warning", message="Inspect warning")],
                markdown="| Quarter | Revenue |\n| --- | --- |\n| Q1 | 10 |\n\n![Revenue chart](s001-b002-chart-001.png)",
                render_result=render_result,
                llm_result=LlmRunResult(
                    status="succeeded",
                    attempts=1,
                    used_model="gpt-5.4",
                    response=LlmResponse(
                        sheet_summary="Summary sheet",
                        sections=[],
                        figures=[],
                        unknowns=["Chart labels are approximate."],
                        markdown="| Quarter | Revenue |\n| --- | --- |\n| Q1 | 10 |\n\n![Revenue chart](s001-b002-chart-001.png)",
                    ),
                ),
            ),
            ConvertSheetResult(
                sheet=failed_sheet,
                status="failed",
                failures=list(failed_sheet.failures),
            ),
        ],
    )

    output_files = write_convert_outputs(convert_result)

    result_markdown = output_files.result_markdown.read_text(encoding="utf-8")
    manifest = json.loads(output_files.manifest_json.read_text(encoding="utf-8"))

    assert "## Summary" in result_markdown
    assert "| Quarter | Revenue |" in result_markdown
    assert "assets/sheet-001/s001-b002-chart-001.png" in result_markdown
    assert "](s001-b002-chart-001.png)" not in result_markdown
    assert "assets/sheet-001/s001-b001-paragraph-range-001.png" not in result_markdown
    assert "## Broken" in result_markdown
    assert "Failed to convert this sheet." in result_markdown
    assert "[formula_cached_value_missing]" in result_markdown

    assert manifest["schema_version"] == "phase1.0"
    assert manifest["input_file_name"] == "book.xlsx"
    assert manifest["command_options"]["model"] == "text-model"
    assert [sheet["status"] for sheet in manifest["sheets"]] == ["succeeded", "failed"]
    assert manifest["sheets"][0]["render"]["status"] == "succeeded"
    assert manifest["sheets"][0]["llm"]["used_model"] == "gpt-5.4"
    assert manifest["sheets"][1]["llm"]["status"] == "skipped"
    assert manifest["blocks"][1]["visual_id"] == "s001-v001-chart"
    assert manifest["blocks"][1]["related_block_id"] == "s001-b001-paragraph"
    assert manifest["blocks"][1]["assets"] == [
        {
            "path": "assets/sheet-001/s001-b002-chart-001.png",
            "kind": "chart",
            "role": "markdown",
            "source": "chart_export",
            "anchor": {
                "sheet": "Summary",
                "start_row": 3,
                "start_col": 1,
                "end_row": 8,
                "end_col": 6,
                "a1": "A3:F8",
            },
        }
    ]
    assert (tmp_path / "out" / "assets" / "sheet-001" / "s001-b002-chart-001.png").is_file()
    assert not (tmp_path / "out" / "assets" / "sheet-001" / "s001-b001-paragraph-range-001.png").exists()


def test_write_convert_outputs_writes_debug_and_optional_render_artifacts(tmp_path: Path) -> None:
    sheet = SheetModel(
        sheet_index=1,
        name="DebugSheet",
        blocks=[
            ParagraphBlock(
                id="s001-b001-paragraph",
                anchor=Rect(sheet="DebugSheet", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                source=SourceKind.CELLS,
                text="Paragraph",
            )
        ],
    )
    render_result = RenderSheetResult(
        input_file_name="book.xlsx",
        sheet_name="DebugSheet",
        temp_dir=str(tmp_path / "render-debug"),
        artifacts=[
            _artifact(
                tmp_path / "render-debug" / "range.png",
                block_id="s001-b001-paragraph",
                kind="range",
                role="render_artifact",
                source="range_copy_picture",
                anchor=Rect(sheet="DebugSheet", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                related_block_id=None,
                visual_id=None,
            )
        ],
    )
    convert_result = ConvertResult(
        input_file_name="book.xlsx",
        schema_version="phase1.0",
        generated_at="2026-04-23T00:00:00+00:00",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": True,
            "save_render_artifacts": True,
            "strict": False,
        },
        output_dir=tmp_path / "out",
        workbook_extraction_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        block_detection_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": [{"name": "DebugSheet"}]},
        linked_workbook_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        sheets=[
            ConvertSheetResult(
                sheet=sheet,
                status="succeeded",
                markdown="Paragraph body",
                render_plan_payload={
                    "sheet_index": 1,
                    "name": "DebugSheet",
                    "items": [{"block_id": "s001-b001-paragraph", "kind": "range", "role": "render_artifact", "source": "range_copy_picture", "target_part": None}],
                    "warnings": [],
                    "failures": [],
                },
                render_result=render_result,
                llm_input_payload={"sheetName": "DebugSheet", "blocks": [], "assets": [], "instructions": {}},
                llm_result=LlmRunResult(
                    status="succeeded",
                    attempts=1,
                    response=LlmResponse(
                        sheet_summary="Debug",
                        sections=[],
                        figures=[],
                        unknowns=[],
                        markdown="Paragraph body",
                    ),
                ),
            )
        ],
    )

    output_files = write_convert_outputs(convert_result)

    assert output_files.debug_dir is not None
    assert (tmp_path / "out" / "assets" / "sheet-001" / "s001-b001-paragraph-range-001.png").is_file()
    assert (output_files.debug_dir / "workbook_extraction.json").is_file()
    assert (output_files.debug_dir / "block_detection.json").is_file()
    assert (output_files.debug_dir / "linked_blocks.json").is_file()
    assert (output_files.debug_dir / "render_plan.json").is_file()
    assert (output_files.debug_dir / "llm_input.json").is_file()
    assert (output_files.debug_dir / "llm_response.json").is_file()
    result_markdown = output_files.result_markdown.read_text(encoding="utf-8")
    assert "assets/sheet-001/s001-b001-paragraph-range-001.png" not in result_markdown


def test_write_convert_outputs_does_not_rewrite_already_published_asset_path(tmp_path: Path) -> None:
    sheet = SheetModel(
        sheet_index=1,
        name="Assets",
        blocks=[
            ChartBlock(
                id="s001-b001-chart",
                anchor=Rect(sheet="Assets", start_row=1, start_col=1, end_row=4, end_col=4, a1="A1:D4"),
                source=SourceKind.CHART,
                title="Chart",
            )
        ],
    )
    render_result = RenderSheetResult(
        input_file_name="book.xlsx",
        sheet_name="Assets",
        temp_dir=str(tmp_path / "render-assets"),
        artifacts=[
            _artifact(
                tmp_path / "render-assets" / "chart.png",
                block_id="s001-b001-chart",
                kind="chart",
                role="markdown",
                source="chart_export",
                anchor=Rect(sheet="Assets", start_row=1, start_col=1, end_row=4, end_col=4, a1="A1:D4"),
                related_block_id=None,
                visual_id=None,
            )
        ],
    )
    convert_result = ConvertResult(
        input_file_name="book.xlsx",
        schema_version="phase1.0",
        generated_at="2026-04-23T00:00:00+00:00",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
        output_dir=tmp_path / "out",
        workbook_extraction_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        block_detection_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        linked_workbook_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        sheets=[
            ConvertSheetResult(
                sheet=sheet,
                status="succeeded",
                markdown="![Chart](assets/sheet-001/s001-b001-chart-001.png)",
                render_result=render_result,
                llm_result=LlmRunResult(
                    status="succeeded",
                    attempts=1,
                    response=LlmResponse(
                        sheet_summary="Assets",
                        sections=[],
                        figures=[],
                        unknowns=[],
                        markdown="![Chart](assets/sheet-001/s001-b001-chart-001.png)",
                    ),
                ),
            )
        ],
    )

    output_files = write_convert_outputs(convert_result)

    result_markdown = output_files.result_markdown.read_text(encoding="utf-8")
    assert "assets/sheet-001/assets/sheet-001" not in result_markdown
    assert "![Chart](assets/sheet-001/s001-b001-chart-001.png)" in result_markdown


def test_write_convert_outputs_redacts_absolute_paths_embedded_in_failure_details(tmp_path: Path) -> None:
    local_path = r"C:\Users\Alice\My Documents\private book.xlsx"
    temp_path = r"C:\Users\Alice\AppData\Local\Temp\excel-semantic-md-render-123\artifact.png"
    sheet = SheetModel(
        sheet_index=1,
        name="PathLeak",
        blocks=[],
    )
    convert_result = ConvertResult(
        input_file_name="book.xlsx",
        schema_version="phase1.0",
        generated_at="2026-04-23T00:00:00+00:00",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
        output_dir=tmp_path / "out",
        workbook_extraction_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        block_detection_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        linked_workbook_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        sheets=[
            ConvertSheetResult(
                sheet=sheet,
                status="failed",
                warnings=[
                    WarningInfo(
                        code="path_warning",
                        message="Warning contains an exception string.",
                        details={"error": f"failed while reading {local_path}", "temp_dir": temp_path},
                    )
                ],
                failures=[
                    FailureInfo(
                        stage="render",
                        message="Render failed.",
                        details={"error": f"Excel failed on {local_path}", "workbook": local_path},
                    )
                ],
            )
        ],
    )

    output_files = write_convert_outputs(convert_result)
    manifest_text = output_files.manifest_json.read_text(encoding="utf-8")
    manifest = json.loads(manifest_text)

    assert local_path not in manifest_text
    assert temp_path not in manifest_text
    assert manifest["sheets"][0]["warnings"][0]["details"]["error"] == "failed while reading [redacted]"
    assert manifest["sheets"][0]["warnings"][0]["details"]["temp_dir"] == "[redacted]"
    assert manifest["sheets"][0]["failures"][0]["details"]["error"] == "Excel failed on [redacted]"
    assert manifest["sheets"][0]["failures"][0]["details"]["workbook"] == "[redacted]"


def test_write_convert_outputs_restores_existing_outputs_when_publish_move_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    output_dir = tmp_path / "out"
    (output_dir / "assets").mkdir(parents=True)
    (output_dir / "debug").mkdir()
    (output_dir / "result.md").write_text("old result\n", encoding="utf-8")
    (output_dir / "manifest.json").write_text('{"old": true}\n', encoding="utf-8")
    (output_dir / "assets" / "old.txt").write_text("old asset\n", encoding="utf-8")
    (output_dir / "debug" / "old.json").write_text('{"old": true}\n', encoding="utf-8")

    sheet = SheetModel(sheet_index=1, name="Restore", blocks=[])
    convert_result = ConvertResult(
        input_file_name="book.xlsx",
        schema_version="phase1.0",
        generated_at="2026-04-23T00:00:00+00:00",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
        output_dir=output_dir,
        workbook_extraction_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        block_detection_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        linked_workbook_payload={"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        sheets=[ConvertSheetResult(sheet=sheet, status="succeeded", markdown="new result")],
    )
    real_move = output_writers.shutil.move

    def guarded_move(src: str, dst: str):
        source_path = Path(src)
        if source_path.name == "manifest.json" and source_path.parent.name.startswith(".excel-semantic-md-staging-"):
            raise RuntimeError("publish move failed")
        return real_move(src, dst)

    monkeypatch.setattr(output_writers.shutil, "move", guarded_move)

    with pytest.raises(RuntimeError, match="publish move failed"):
        write_convert_outputs(convert_result)

    assert (output_dir / "result.md").read_text(encoding="utf-8") == "old result\n"
    assert (output_dir / "manifest.json").read_text(encoding="utf-8") == '{"old": true}\n'
    assert (output_dir / "assets" / "old.txt").read_text(encoding="utf-8") == "old asset\n"
    assert (output_dir / "debug" / "old.json").read_text(encoding="utf-8") == '{"old": true}\n'


def test_convert_command_writes_outputs_and_only_fails_in_strict_mode(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    input_path = tmp_path / "book.xlsx"
    input_path.write_bytes(b"placeholder")
    base_out = tmp_path / "convert-out"

    def fake_read_workbook(path: Path):
        return SimpleNamespace(
            input_file_name=path.name,
            to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": path.name, "sheets": []},
        )

    success_sheet = SheetModel(
        sheet_index=1,
        name="Success",
        blocks=[
            ParagraphBlock(
                id="s001-b001-paragraph",
                anchor=Rect(sheet="Success", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                source=SourceKind.CELLS,
                text="Success paragraph",
            )
        ],
    )
    failed_sheet = SheetModel(
        sheet_index=2,
        name="Failed",
        blocks=[],
        failures=[FailureInfo(stage="workbook_reading", message="Formula cache missing.", details={"code": "formula_cached_value_missing"})],
    )

    def fake_detect_blocks(_read_result):
        return SimpleNamespace(
            schema_version="phase1.0",
            input_file_name="book.xlsx",
            sheets=[success_sheet, failed_sheet],
            to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
        )

    def fake_read_visual_metadata(_path: Path):
        return []

    def fake_link_visuals(block_model, _visuals):
        return block_model

    def fake_build_render_plan(sheet, _visual_sheet, *, save_render_artifacts: bool):
        return (
            [
                RenderPlanItem(
                    block=sheet.blocks[0],
                    kind="range",
                    role=SimpleNamespace(value="render_artifact"),
                    source="range_copy_picture",
                )
            ],
            [],
            [],
        )

    def fake_render_with_excel_com(
        _input_path: Path,
        *,
        input_file_name: str,
        sheet_name: str,
        plan_items,
        warnings,
        failures,
    ):
        temp_dir = tmp_path / f"render-{sheet_name}"
        artifact_path = temp_dir / "sheet-001" / "range.png"
        artifact_path.parent.mkdir(parents=True, exist_ok=True)
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

    class FakeAdapter:
        def run_sheet(self, sheet, render_result, *, options, request):
            assert render_result is None
            assert options.model == "text-model"
            assert request.input.sheet_name == sheet.name
            assert request.attachments == []
            return LlmRunResult(
                status="succeeded",
                attempts=1,
                response=LlmResponse(
                    sheet_summary=f"{sheet.name} summary",
                    sections=[],
                    figures=[],
                    unknowns=[],
                    markdown=f"Markdown for {sheet.name}",
                ),
            )

    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_workbook", fake_read_workbook)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.detect_blocks", fake_detect_blocks)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_visual_metadata", fake_read_visual_metadata)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.link_visuals", fake_link_visuals)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.build_render_plan", fake_build_render_plan)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.render_with_excel_com", fake_render_with_excel_com)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.GitHubCopilotSdkAdapter", FakeAdapter)

    code = cli_main.main(
        [
            "convert",
            "--input",
            str(input_path),
            "--out",
            str(base_out),
            "--model",
            "text-model",
        ]
    )

    manifest = json.loads((base_out / "manifest.json").read_text(encoding="utf-8"))
    result_markdown = (base_out / "result.md").read_text(encoding="utf-8")

    assert code == 0
    assert [sheet["status"] for sheet in manifest["sheets"]] == ["succeeded", "failed"]
    assert "## Success" in result_markdown
    assert "Markdown for Success" in result_markdown
    assert "## Failed" in result_markdown
    assert "Failed to convert this sheet." in result_markdown

    strict_out = tmp_path / "convert-out-strict"
    code = cli_main.main(
        [
            "convert",
            "--input",
            str(input_path),
            "--out",
            str(strict_out),
            "--model",
            "text-model",
            "--strict",
        ]
    )

    assert code == 1
    assert (strict_out / "result.md").is_file()
    assert (strict_out / "manifest.json").is_file()


def test_run_convert_pipeline_uses_prepared_llm_request_for_debug_payload(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    input_path = tmp_path / "book.xlsx"
    input_path.write_bytes(b"placeholder")
    render_path = tmp_path / "render-prepared-request" / "sheet-001" / "chart.png"
    render_path.parent.mkdir(parents=True)
    render_path.write_bytes(b"png")
    observed: dict[str, object] = {}
    sheet = SheetModel(
        sheet_index=1,
        name="PreparedRequest",
        blocks=[
            ChartBlock(
                id="s001-b001-chart",
                anchor=Rect(sheet="PreparedRequest", start_row=2, start_col=2, end_row=6, end_col=6, a1="B2:F6"),
                source=SourceKind.CHART,
                visual_id="s001-v001-chart",
                title="Prepared Chart",
            )
        ],
    )
    read_result = SimpleNamespace(
        input_file_name="book.xlsx",
        to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
    )
    linked_workbook = SimpleNamespace(
        schema_version="phase1.0",
        input_file_name="book.xlsx",
        sheets=[sheet],
        to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
    )
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_workbook", lambda _path: read_result)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.detect_blocks", lambda _read_result: linked_workbook)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_visual_metadata", lambda _path: [])
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.link_visuals", lambda block_model, _visuals: block_model)
    monkeypatch.setattr(
        "excel_semantic_md.app.convert_pipeline.build_render_plan",
        lambda _sheet, _visual_sheet, *, save_render_artifacts: (
            [
                RenderPlanItem(
                    block=sheet.blocks[0],
                    kind="chart",
                    role=SimpleNamespace(value="markdown"),
                    source="chart_export",
                )
            ],
            [],
            [],
        ),
    )
    monkeypatch.setattr(
        "excel_semantic_md.app.convert_pipeline.render_with_excel_com",
        lambda *args, **kwargs: RenderSheetResult(
            input_file_name="book.xlsx",
            sheet_name="PreparedRequest",
            temp_dir=str(tmp_path / "render-prepared-request"),
            artifacts=[
                RenderArtifact(
                    block_id=sheet.blocks[0].id,
                    visual_id="s001-v001-chart",
                    related_block_id=None,
                    kind="chart",
                    role="markdown",
                    path=str(render_path),
                    source="chart_export",
                    anchor=sheet.blocks[0].anchor,
                )
            ],
            warnings=[],
            failures=[],
        ),
    )

    class ObservingAdapter:
        def run_sheet(self, _sheet, _render_result, *, options, request):
            observed["options"] = options
            observed["request_input"] = request.input.to_dict()
            observed["request_attachments"] = [attachment.to_dict() for attachment in request.attachments]
            observed["request_prompt"] = request.prompt
            return LlmRunResult(
                status="succeeded",
                attempts=1,
                response=LlmResponse(
                    sheet_summary="Prepared",
                    sections=[],
                    figures=[],
                    unknowns=[],
                    markdown="Prepared markdown",
                ),
            )

    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.GitHubCopilotSdkAdapter", ObservingAdapter)

    result = run_convert_pipeline(
        input_path,
        tmp_path / "out",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": 1,
            "save_debug_json": True,
            "save_render_artifacts": False,
            "strict": False,
        },
    )
    output_files = write_convert_outputs(result)
    debug_payload = json.loads((output_files.debug_dir / "llm_input.json").read_text(encoding="utf-8"))

    assert result.sheets[0].llm_input_payload == observed["request_input"]
    assert debug_payload["sheets"][0]["input"] == observed["request_input"]
    assert observed["request_attachments"] == [
        {
            "path": str(render_path.resolve()),
            "block_id": "s001-b001-chart",
            "related_block_id": None,
            "kind": "chart",
            "source": "chart_export",
            "priority": 0,
        }
    ]
    assert debug_payload["sheets"][0]["input"]["assets"] == [
        {
            "path": "chart.png",
            "block_id": "s001-b001-chart",
            "related_block_id": None,
            "kind": "chart",
            "source": "chart_export",
            "priority": 0,
        }
    ]
    assert "PreparedRequest" in observed["request_prompt"]


def test_run_convert_pipeline_normalizes_render_exception_to_failed_sheet(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    input_path = tmp_path / "book.xlsx"
    input_path.write_bytes(b"placeholder")
    sheet = SheetModel(
        sheet_index=1,
        name="RenderExplodes",
        blocks=[
            ChartBlock(
                id="s001-b001-chart",
                anchor=Rect(sheet="RenderExplodes", start_row=2, start_col=2, end_row=6, end_col=6, a1="B2:F6"),
                source=SourceKind.CHART,
                visual_id="s001-v001-chart",
                title="Exploding chart",
            )
        ],
    )
    read_result = SimpleNamespace(
        input_file_name="book.xlsx",
        to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
    )
    linked_workbook = SimpleNamespace(
        schema_version="phase1.0",
        input_file_name="book.xlsx",
        sheets=[sheet],
        to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
    )
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_workbook", lambda _path: read_result)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.detect_blocks", lambda _read_result: linked_workbook)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_visual_metadata", lambda _path: [])
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.link_visuals", lambda block_model, _visuals: block_model)
    monkeypatch.setattr(
        "excel_semantic_md.app.convert_pipeline.build_render_plan",
        lambda _sheet, _visual_sheet, *, save_render_artifacts: (
            [
                RenderPlanItem(
                    block=sheet.blocks[0],
                    kind="chart",
                    role=SimpleNamespace(value="markdown"),
                    source="chart_export",
                )
            ],
            [],
            [],
        ),
    )
    monkeypatch.setattr(
        "excel_semantic_md.app.convert_pipeline.render_with_excel_com",
        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("render boom")),
    )
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.GitHubCopilotSdkAdapter", lambda: SimpleNamespace())

    result = run_convert_pipeline(
        input_path,
        tmp_path / "out",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
    )

    assert result.failed_sheet_count == 1
    assert result.sheets[0].status == "failed"
    assert result.sheets[0].failures[0].stage == "render"
    assert "render boom" in result.sheets[0].failures[0].details["error"]


def test_run_convert_pipeline_normalizes_render_plan_exception_to_failed_sheet(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_path = tmp_path / "book.xlsx"
    input_path.write_bytes(b"placeholder")
    success_sheet = SheetModel(
        sheet_index=1,
        name="Success",
        blocks=[],
    )
    broken_sheet = SheetModel(
        sheet_index=2,
        name="PlanExplodes",
        blocks=[
            ParagraphBlock(
                id="s002-b001-paragraph",
                anchor=Rect(sheet="PlanExplodes", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                source=SourceKind.CELLS,
                text="Paragraph",
            )
        ],
    )
    read_result = SimpleNamespace(
        input_file_name="book.xlsx",
        to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
    )
    linked_workbook = SimpleNamespace(
        schema_version="phase1.0",
        input_file_name="book.xlsx",
        sheets=[success_sheet, broken_sheet],
        to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
    )

    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_workbook", lambda _path: read_result)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.detect_blocks", lambda _read_result: linked_workbook)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_visual_metadata", lambda _path: [])
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.link_visuals", lambda block_model, _visuals: block_model)

    def fake_build_render_plan(sheet, _visual_sheet, *, save_render_artifacts: bool):
        if sheet.name == "PlanExplodes":
            raise RuntimeError("plan boom")
        return ([], [], [])

    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.build_render_plan", fake_build_render_plan)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.GitHubCopilotSdkAdapter", lambda: SimpleNamespace())

    result = run_convert_pipeline(
        input_path,
        tmp_path / "out",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
    )

    assert [sheet.status for sheet in result.sheets] == ["succeeded", "failed"]
    assert result.sheets[0].llm_result is not None
    assert result.sheets[0].llm_result.status == "succeeded"
    assert result.sheets[1].failures[0].stage == "render_plan"
    assert "plan boom" in result.sheets[1].failures[0].details["error"]

    output_files = write_convert_outputs(result)
    manifest = json.loads(output_files.manifest_json.read_text(encoding="utf-8"))
    result_markdown = output_files.result_markdown.read_text(encoding="utf-8")
    assert [sheet["status"] for sheet in manifest["sheets"]] == ["succeeded", "failed"]
    assert manifest["sheets"][1]["render"]["status"] == "failed"
    assert manifest["sheets"][1]["render"]["failures"][0]["stage"] == "render_plan"
    assert manifest["sheets"][1]["llm"]["status"] == "skipped"
    assert "## PlanExplodes" in result_markdown
    assert "Failed to convert this sheet." in result_markdown


def test_run_convert_pipeline_treats_failed_llm_status_without_failure_object_as_failed_sheet(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_path = tmp_path / "book.xlsx"
    input_path.write_bytes(b"placeholder")
    sheet = SheetModel(
        sheet_index=1,
        name="LlmFails",
        blocks=[
            ParagraphBlock(
                id="s001-b001-paragraph",
                anchor=Rect(sheet="LlmFails", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                source=SourceKind.CELLS,
                text="Paragraph",
            )
        ],
    )
    render_path = tmp_path / "render-llm-fails" / "sheet-001" / "range.png"
    render_path.parent.mkdir(parents=True, exist_ok=True)
    render_path.write_bytes(b"png")
    read_result = SimpleNamespace(
        input_file_name="book.xlsx",
        to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
    )
    linked_workbook = SimpleNamespace(
        schema_version="phase1.0",
        input_file_name="book.xlsx",
        sheets=[sheet],
        to_dict=lambda: {"schema_version": "phase1.0", "input_file_name": "book.xlsx", "sheets": []},
    )
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_workbook", lambda _path: read_result)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.detect_blocks", lambda _read_result: linked_workbook)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.read_visual_metadata", lambda _path: [])
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.link_visuals", lambda block_model, _visuals: block_model)
    monkeypatch.setattr(
        "excel_semantic_md.app.convert_pipeline.build_render_plan",
        lambda _sheet, _visual_sheet, *, save_render_artifacts: (
            [
                RenderPlanItem(
                    block=sheet.blocks[0],
                    kind="range",
                    role=SimpleNamespace(value="render_artifact"),
                    source="range_copy_picture",
                )
            ],
            [],
            [],
        ),
    )
    monkeypatch.setattr(
        "excel_semantic_md.app.convert_pipeline.render_with_excel_com",
        lambda *args, **kwargs: RenderSheetResult(
            input_file_name="book.xlsx",
            sheet_name="LlmFails",
            temp_dir=str(tmp_path / "render-llm-fails"),
            artifacts=[
                RenderArtifact(
                    block_id=sheet.blocks[0].id,
                    visual_id=None,
                    related_block_id=None,
                    kind="range",
                    role="render_artifact",
                    path=str(render_path),
                    source="range_copy_picture",
                    anchor=sheet.blocks[0].anchor,
                )
            ],
            warnings=[],
            failures=[],
        ),
    )

    class FailedAdapter:
        def run_sheet(self, *_args, **_kwargs):
            return LlmRunResult(status="failed", attempts=1, response=None, failure=None)

    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.GitHubCopilotSdkAdapter", FailedAdapter)

    result = run_convert_pipeline(
        input_path,
        tmp_path / "out",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
    )

    assert result.sheets[0].status == "failed"
    assert result.sheets[0].failures[0].stage == "llm"
    assert "without details" in result.sheets[0].failures[0].message


def test_run_convert_pipeline_skips_render_for_cell_only_sheet(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    input_path = Path(__file__).resolve().parent / "fixtures" / "visuals" / "no-visuals.xlsx"
    observed: dict[str, object] = {}

    class ObservingAdapter:
        def run_sheet(self, sheet, render_result, *, options, request):
            observed["sheet_name"] = sheet.name
            observed["render_result"] = render_result
            observed["attachments"] = [attachment.to_dict() for attachment in request.attachments]
            return LlmRunResult(
                status="succeeded",
                attempts=1,
                response=LlmResponse(
                    sheet_summary="Plain sheet",
                    sections=[],
                    figures=[],
                    unknowns=[],
                    markdown="Plain markdown",
                ),
            )

    monkeypatch.setattr(
        "excel_semantic_md.app.convert_pipeline.render_with_excel_com",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("render must not run for cell-only sheets")),
    )
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.GitHubCopilotSdkAdapter", ObservingAdapter)

    result = run_convert_pipeline(
        input_path,
        tmp_path / "out",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
    )

    assert result.failed_sheet_count == 0
    assert result.sheets[0].status == "succeeded"
    assert result.sheets[0].render_result is None
    assert result.sheets[0].markdown == "Plain markdown"
    assert observed == {
        "sheet_name": "Plain",
        "render_result": None,
        "attachments": [],
    }


def test_run_convert_pipeline_skips_render_for_cell_only_sheet_with_zero_image_limit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    input_path = Path(__file__).resolve().parent / "fixtures" / "visuals" / "no-visuals.xlsx"
    observed: dict[str, object] = {}

    class ObservingAdapter:
        def run_sheet(self, _sheet, render_result, *, options, request):
            observed["render_result"] = render_result
            observed["max_images_per_sheet"] = options.max_images_per_sheet
            observed["attachments"] = [attachment.to_dict() for attachment in request.attachments]
            return LlmRunResult(
                status="succeeded",
                attempts=1,
                response=LlmResponse(
                    sheet_summary="Plain sheet",
                    sections=[],
                    figures=[],
                    unknowns=[],
                    markdown="Plain markdown",
                ),
            )

    monkeypatch.setattr(
        "excel_semantic_md.app.convert_pipeline.render_with_excel_com",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("render must not run when image limit is zero and no visual blocks exist")),
    )
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.GitHubCopilotSdkAdapter", ObservingAdapter)

    result = run_convert_pipeline(
        input_path,
        tmp_path / "out",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": 0,
            "save_debug_json": False,
            "save_render_artifacts": False,
            "strict": False,
        },
    )

    assert result.failed_sheet_count == 0
    assert result.sheets[0].status == "succeeded"
    assert observed == {
        "render_result": None,
        "max_images_per_sheet": 0,
        "attachments": [],
    }


def test_run_convert_pipeline_preserves_explicit_render_artifacts_for_cell_only_sheet(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    input_path = Path(__file__).resolve().parent / "fixtures" / "visuals" / "no-visuals.xlsx"
    render_path = tmp_path / "render-cell-only" / "sheet-001" / "range.png"
    observed: dict[str, object] = {}

    class ObservingAdapter:
        def run_sheet(self, _sheet, render_result, *, options, request):
            observed["render_result"] = render_result
            observed["attachments"] = [attachment.to_dict() for attachment in request.attachments]
            observed["save_render_artifacts"] = options.max_images_per_sheet
            return LlmRunResult(
                status="succeeded",
                attempts=1,
                response=LlmResponse(
                    sheet_summary="Plain sheet",
                    sections=[],
                    figures=[],
                    unknowns=[],
                    markdown="Plain markdown",
                ),
            )

    def fake_render_with_excel_com(
        _input_path: Path,
        *,
        input_file_name: str,
        sheet_name: str,
        plan_items,
        warnings,
        failures,
    ) -> RenderSheetResult:
        render_path.parent.mkdir(parents=True, exist_ok=True)
        render_path.write_bytes(b"png")
        return RenderSheetResult(
            input_file_name=input_file_name,
            sheet_name=sheet_name,
            temp_dir=str(tmp_path / "render-cell-only"),
            artifacts=[
                RenderArtifact(
                    block_id=plan_items[0].block.id,
                    visual_id=None,
                    related_block_id=None,
                    kind="range",
                    role="render_artifact",
                    path=str(render_path),
                    source="range_copy_picture",
                    anchor=plan_items[0].block.anchor,
                )
            ],
            warnings=list(warnings),
            failures=list(failures),
        )

    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.render_with_excel_com", fake_render_with_excel_com)
    monkeypatch.setattr("excel_semantic_md.app.convert_pipeline.GitHubCopilotSdkAdapter", ObservingAdapter)

    result = run_convert_pipeline(
        input_path,
        tmp_path / "out",
        command_options={
            "model": None,
            "vision_model": None,
            "max_images_per_sheet": None,
            "save_debug_json": False,
            "save_render_artifacts": True,
            "strict": False,
        },
    )

    assert result.failed_sheet_count == 0
    assert result.sheets[0].render_result is not None
    assert observed["render_result"] is not None
    assert observed["attachments"] == []


def _artifact(
    path: Path,
    *,
    block_id: str,
    kind: str,
    role: str,
    source: str,
    anchor: Rect,
    related_block_id: str | None,
    visual_id: str | None,
) -> RenderArtifact:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"png")
    return RenderArtifact(
        block_id=block_id,
        visual_id=visual_id,
        related_block_id=related_block_id,
        kind=kind,
        role=role,
        path=str(path),
        source=source,
        anchor=anchor,
    )
