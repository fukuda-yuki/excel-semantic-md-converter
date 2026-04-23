"""Convert orchestration for Phase 1 output generation."""

from __future__ import annotations

import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from excel_semantic_md.excel import detect_blocks, link_visuals, read_visual_metadata, read_workbook
from excel_semantic_md.excel.ooxml_visual_reader import SheetVisualResult
from excel_semantic_md.llm import GitHubCopilotSdkAdapter, LlmRunOptions, build_llm_request
from excel_semantic_md.llm.models import LlmResponse, LlmRunResult
from excel_semantic_md.models import FailureInfo, SheetModel, WarningInfo
from excel_semantic_md.output.models import ConvertResult, ConvertSheetResult
from excel_semantic_md.render import build_render_plan, render_with_excel_com


def run_convert_pipeline(
    input_path: Path,
    output_dir: Path,
    *,
    command_options: dict[str, Any],
) -> ConvertResult:
    workbook_read = read_workbook(input_path)
    block_model = detect_blocks(workbook_read)
    visual_results = read_visual_metadata(input_path)
    linked_workbook = link_visuals(block_model, visual_results)
    visual_results_by_index = {sheet.sheet_index: sheet for sheet in visual_results}
    llm_adapter = GitHubCopilotSdkAdapter()

    sheet_results: list[ConvertSheetResult] = []
    try:
        for linked_sheet in linked_workbook.sheets:
            visual_sheet = visual_results_by_index.get(linked_sheet.sheet_index)
            sheet_results.append(
                _run_sheet_pipeline(
                    input_path=input_path,
                    linked_sheet=linked_sheet,
                    visual_sheet=visual_sheet,
                    input_file_name=workbook_read.input_file_name,
                    command_options=command_options,
                    llm_adapter=llm_adapter,
                )
            )
    except Exception:
        cleanup_convert_result(
            ConvertResult(
                input_file_name=workbook_read.input_file_name,
                schema_version=linked_workbook.schema_version,
                generated_at=datetime.now(timezone.utc).isoformat(),
                command_options=dict(command_options),
                output_dir=output_dir,
                workbook_extraction_payload=workbook_read.to_dict(),
                block_detection_payload=block_model.to_dict(),
                linked_workbook_payload=linked_workbook.to_dict(),
                sheets=sheet_results,
            )
        )
        raise

    return ConvertResult(
        input_file_name=workbook_read.input_file_name,
        schema_version=linked_workbook.schema_version,
        generated_at=datetime.now(timezone.utc).isoformat(),
        command_options=dict(command_options),
        output_dir=output_dir,
        workbook_extraction_payload=workbook_read.to_dict(),
        block_detection_payload=block_model.to_dict(),
        linked_workbook_payload=linked_workbook.to_dict(),
        sheets=sheet_results,
    )


def cleanup_convert_result(result: ConvertResult) -> list[WarningInfo]:
    seen_dirs: set[str] = set()
    cleanup_errors: list[WarningInfo] = []
    for sheet_result in result.sheets:
        render_result = sheet_result.render_result
        if render_result is None or render_result.temp_dir in seen_dirs:
            continue
        seen_dirs.add(render_result.temp_dir)
        try:
            shutil.rmtree(render_result.temp_dir)
        except OSError as exc:
            cleanup_errors.append(
                WarningInfo(
                    code="render_temp_cleanup_failed",
                    message="Render temp directory cleanup failed after convert.",
                    details={"temp_dir": "[redacted]", "error": str(exc)},
                )
            )
    return cleanup_errors


def _run_sheet_pipeline(
    *,
    input_path: Path,
    linked_sheet: SheetModel,
    visual_sheet: SheetVisualResult | None,
    input_file_name: str,
    command_options: dict[str, Any],
    llm_adapter: GitHubCopilotSdkAdapter,
) -> ConvertSheetResult:
    warnings = list(linked_sheet.warnings)
    failures = list(linked_sheet.failures)
    if visual_sheet is not None:
        warnings.extend(_warning_info(item) for item in visual_sheet.warnings)

    render_plan_payload: dict[str, Any] | None = None
    render_result = None
    llm_input_payload: dict[str, Any] | None = None
    llm_result: LlmRunResult | None = None
    markdown: str | None = None
    stage = "sheet"

    try:
        if not failures:
            stage = "render_plan"
            plan_items, plan_warnings, plan_failures = build_render_plan(
                linked_sheet,
                visual_sheet,
                save_render_artifacts=bool(command_options.get("save_render_artifacts")),
            )
            render_plan_payload = {
                "sheet_index": linked_sheet.sheet_index,
                "name": linked_sheet.name,
                "items": [
                    {
                        "block_id": item.block.id,
                        "kind": item.kind,
                        "role": item.role.value,
                        "source": item.source,
                        "target_part": item.target_part,
                    }
                    for item in plan_items
                ],
                "warnings": [warning.to_dict() for warning in plan_warnings],
                "failures": [failure.to_dict() for failure in plan_failures],
            }
            warnings.extend(plan_warnings)
            failures.extend(plan_failures)

            if not failures and not plan_items and not linked_sheet.blocks:
                stage = "llm_input"
                llm_options = LlmRunOptions(
                    model=command_options.get("model"),
                    vision_model=command_options.get("vision_model"),
                    max_images_per_sheet=command_options.get("max_images_per_sheet"),
                )
                llm_request = build_llm_request(linked_sheet, None, options=llm_options)
                llm_input_payload = llm_request.input.to_dict()
                llm_result = LlmRunResult(
                    status="succeeded",
                    attempts=1,
                    response=LlmResponse(
                        sheet_summary="No visible content.",
                        sections=[],
                        figures=[],
                        unknowns=[],
                        markdown="",
                        raw={"generated_by": "empty_sheet_short_circuit"},
                    ),
                )
                markdown = ""
            elif not failures:
                try:
                    stage = "render"
                    render_result = render_with_excel_com(
                        input_path,
                        input_file_name=input_file_name,
                        sheet_name=linked_sheet.name,
                        plan_items=plan_items,
                        warnings=[],
                        failures=[],
                    )
                except Exception as exc:
                    failures.append(
                        FailureInfo(
                            stage="render",
                            message="Render stage raised an unexpected exception.",
                            details={"sheet_name": linked_sheet.name, "error": str(exc)},
                        )
                    )
                else:
                    warnings.extend(render_result.warnings)
                    failures.extend(render_result.failures)

                if not failures:
                    stage = "llm_input"
                    llm_options = LlmRunOptions(
                        model=command_options.get("model"),
                        vision_model=command_options.get("vision_model"),
                        max_images_per_sheet=command_options.get("max_images_per_sheet"),
                    )
                    llm_request = build_llm_request(linked_sheet, render_result, options=llm_options)
                    llm_input_payload = llm_request.input.to_dict()
                    try:
                        stage = "llm"
                        llm_result = llm_adapter.run_sheet(
                            linked_sheet,
                            render_result,
                            options=llm_options,
                            request=llm_request,
                        )
                    except Exception as exc:
                        failures.append(
                            FailureInfo(
                                stage="llm",
                                message="LLM stage raised an unexpected exception.",
                                details={"sheet_name": linked_sheet.name, "error": str(exc)},
                            )
                        )
                    else:
                        if llm_result.status == "failed":
                            failures.append(
                                llm_result.failure
                                or FailureInfo(
                                    stage="llm",
                                    message="LLM stage reported failure without details.",
                                    details={"sheet_name": linked_sheet.name},
                                )
                            )
                        elif llm_result.response is not None:
                            markdown = llm_result.response.markdown
    except Exception as exc:
        failures.append(
            FailureInfo(
                stage=stage,
                message="Sheet pipeline raised an unexpected exception.",
                details={"sheet_name": linked_sheet.name, "error": str(exc)},
            )
        )

    status = "failed" if failures else "succeeded"
    return ConvertSheetResult(
        sheet=linked_sheet,
        status=status,
        warnings=warnings,
        failures=failures,
        markdown=markdown,
        render_plan_payload=render_plan_payload,
        render_result=render_result,
        llm_input_payload=llm_input_payload,
        llm_result=llm_result,
    )


def _warning_info(item: Any) -> WarningInfo:
    if isinstance(item, WarningInfo):
        return item
    return WarningInfo(
        code=getattr(item, "code"),
        message=getattr(item, "message"),
        details=dict(getattr(item, "details", {})),
    )
