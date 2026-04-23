"""Builders for sheet-level LLM inputs and attachments."""

from __future__ import annotations

from pathlib import Path

from excel_semantic_md.models import SheetModel
from excel_semantic_md.render.types import RenderArtifact, RenderSheetResult

from excel_semantic_md.llm.models import LlmAttachment, LlmInput, LlmRequest, LlmRunOptions
from excel_semantic_md.llm.prompt import build_sheet_prompt


def build_llm_attachments(
    sheet: SheetModel,
    render_result: RenderSheetResult | None,
    *,
    max_images_per_sheet: int | None,
) -> list[LlmAttachment]:
    if render_result is None:
        return []
    if max_images_per_sheet == 0:
        return []

    blocks_by_id = {block.id: block for block in sheet.blocks}
    ranked = sorted(
        (_artifact_to_attachment(artifact) for artifact in render_result.artifacts),
        key=lambda item: _attachment_sort_key(item, blocks_by_id),
    )
    if max_images_per_sheet is None:
        return ranked
    return ranked[:max_images_per_sheet]


def build_llm_input(sheet: SheetModel, attachments: list[LlmAttachment]) -> LlmInput:
    return LlmInput(
        sheet_name=sheet.name,
        blocks=[block.to_dict() for block in sheet.blocks],
        assets=[_attachment_input_dict(attachment) for attachment in attachments],
        instructions={
            "targetFormat": "markdown",
            "style": "semantic",
            "preserveUnknowns": True,
        },
    )


def build_llm_request(
    sheet: SheetModel,
    render_result: RenderSheetResult | None,
    *,
    options: LlmRunOptions | None = None,
) -> LlmRequest:
    run_options = options or LlmRunOptions()
    attachments = build_llm_attachments(
        sheet,
        render_result,
        max_images_per_sheet=run_options.max_images_per_sheet,
    )
    llm_input = build_llm_input(sheet, attachments)
    return LlmRequest(
        attachments=attachments,
        input=llm_input,
        prompt=build_sheet_prompt(llm_input),
    )


def _artifact_to_attachment(artifact: RenderArtifact) -> LlmAttachment:
    return LlmAttachment(
        path=str(Path(artifact.path).resolve()),
        block_id=artifact.block_id,
        related_block_id=artifact.related_block_id,
        kind=artifact.kind,
        source=artifact.source,
        priority=_attachment_priority(artifact),
    )


def _attachment_priority(artifact: RenderArtifact) -> int:
    if artifact.role == "markdown" and artifact.kind in {"chart", "image", "shape"}:
        return 0
    if artifact.related_block_id is not None:
        return 1
    if artifact.kind == "range" and artifact.role == "render_artifact":
        return 2
    return 3


def _attachment_sort_key(attachment: LlmAttachment, blocks_by_id: dict[str, object]) -> tuple[int, int, str, str]:
    distance = _attachment_distance(attachment, blocks_by_id)
    return (attachment.priority, distance, attachment.kind, attachment.path)


def _attachment_distance(attachment: LlmAttachment, blocks_by_id: dict[str, object]) -> int:
    target_id = attachment.related_block_id or attachment.block_id
    target_block = blocks_by_id.get(target_id) if target_id is not None else None
    if target_block is None:
        return 10**9
    attachment_block = blocks_by_id.get(attachment.block_id) if attachment.block_id is not None else None
    if attachment_block is not None:
        source_anchor = attachment_block.anchor
    else:
        return 10**9
    target_anchor = target_block.anchor
    return abs(source_anchor.start_row - target_anchor.start_row) + abs(source_anchor.start_col - target_anchor.start_col)


def _attachment_input_dict(attachment: LlmAttachment) -> dict[str, object]:
    if hasattr(attachment, "to_dict") and not hasattr(attachment, "path"):
        data = dict(attachment.to_dict())
        data["path"] = Path(str(data["path"])).name
        return data
    return {
        "path": Path(attachment.path).name,
        "block_id": attachment.block_id,
        "related_block_id": attachment.related_block_id,
        "kind": attachment.kind,
        "source": attachment.source,
        "priority": attachment.priority,
    }
