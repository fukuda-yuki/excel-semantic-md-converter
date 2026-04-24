"""Build render plans for Excel COM confirmation."""

from __future__ import annotations

from pathlib import PurePosixPath

from excel_semantic_md.excel.ooxml_visual_reader import SheetVisualResult, VisualElement
from excel_semantic_md.models import AssetRole, Block, FailureInfo, SheetModel, SourceKind, WarningInfo
from excel_semantic_md.render.types import RenderPlanItem


_SAFE_IMAGE_CONTENT_TYPES_BY_EXTENSION = {
    ".png": {"image/png"},
    ".jpg": {"image/jpeg"},
    ".jpeg": {"image/jpeg"},
    ".gif": {"image/gif"},
    ".bmp": {"image/bmp"},
    ".tif": {"image/tiff"},
    ".tiff": {"image/tiff"},
}


def build_render_plan(
    sheet: SheetModel,
    visual_sheet: SheetVisualResult | None,
    *,
    save_render_artifacts: bool = False,
) -> tuple[list[RenderPlanItem], list[WarningInfo], list[FailureInfo]]:
    """Return plan items for a single linked sheet."""

    visuals_by_id = {}
    if visual_sheet is not None:
        visuals_by_id = {visual.id: visual for visual in visual_sheet.visuals}

    items: list[RenderPlanItem] = []
    warnings: list[WarningInfo] = []
    failures: list[FailureInfo] = []

    for block in sheet.blocks:
        if block.source == SourceKind.CELLS:
            items.append(
                RenderPlanItem(
                    block=block,
                    kind="range",
                    role=AssetRole.RENDER_ARTIFACT,
                    source="range_copy_picture",
                )
            )
            continue

        if block.source == SourceKind.CHART:
            items.append(
                RenderPlanItem(
                    block=block,
                    kind="chart",
                    role=AssetRole.MARKDOWN,
                    source="chart_export",
                )
            )
            if save_render_artifacts:
                items.append(
                    RenderPlanItem(
                        block=block,
                        kind="range",
                        role=AssetRole.RENDER_ARTIFACT,
                        source="range_copy_picture",
                    )
                )
            continue

        if block.source in {SourceKind.SHAPE, SourceKind.IMAGE}:
            if block.source == SourceKind.IMAGE:
                visual = visuals_by_id.get(block.visual_id) if block.visual_id is not None else None
                target_part, content_type, warning_code, warning_message = _image_target_part(visual)
                if target_part is None:
                    warnings.append(
                        WarningInfo(
                            code=warning_code,
                            message=warning_message,
                            details={
                                "block_id": block.id,
                                "visual_id": block.visual_id,
                                "content_type": content_type,
                            },
                        )
                    )
                else:
                    items.append(
                        RenderPlanItem(
                            block=block,
                            kind="image",
                            role=AssetRole.MARKDOWN,
                            source="ooxml_image_copy",
                            target_part=target_part,
                        )
                    )
                items.append(
                    RenderPlanItem(
                        block=block,
                        kind="image",
                        role=AssetRole.RENDER_ARTIFACT,
                        source="shape_copy_picture",
                    )
                )
            else:
                items.append(
                    RenderPlanItem(
                        block=block,
                        kind="shape",
                        role=AssetRole.MARKDOWN,
                        source="shape_copy_picture",
                    )
                )
            if save_render_artifacts:
                items.append(
                    RenderPlanItem(
                        block=block,
                        kind="range",
                        role=AssetRole.RENDER_ARTIFACT,
                        source="range_copy_picture",
                    )
                )
            continue

        failures.append(
            FailureInfo(
                stage="render_planning",
                message="Block source is not renderable in the Excel COM milestone.",
                details={"block_id": block.id, "source": block.source.value},
            )
        )

    return items, warnings, failures


def _image_target_part(visual: VisualElement | None) -> tuple[str | None, str | None, str, str]:
    if visual is None or visual.kind != "image":
        return (
            None,
            None,
            "image_original_asset_unavailable",
            "The linked image does not expose a usable OOXML image part, so the original asset copy was skipped.",
        )
    if any(warning.code in {"image_target_missing", "image_part_missing"} for warning in visual.warnings):
        return (
            None,
            visual.asset_candidate.content_type,
            "image_original_asset_unavailable",
            "The linked image does not expose a usable OOXML image part, so the original asset copy was skipped.",
        )
    content_type = visual.asset_candidate.content_type
    if content_type is None or not content_type.startswith("image/"):
        return (
            None,
            content_type,
            "image_original_asset_invalid_content_type",
            "The linked image target is not an image content type, so the original asset copy was skipped.",
        )
    target_part = visual.source.target_part
    extension = PurePosixPath(target_part or "").suffix.lower()
    allowed_content_types = _SAFE_IMAGE_CONTENT_TYPES_BY_EXTENSION.get(extension)
    if target_part is None or not target_part.startswith("xl/media/") or allowed_content_types is None or content_type not in allowed_content_types:
        return (
            None,
            content_type,
            "image_original_asset_untrusted_part",
            "The linked image target is not a trusted OOXML media image part, so the original asset copy was skipped.",
        )
    return (
        target_part,
        content_type,
        "image_original_asset_unavailable",
        "The linked image does not expose a usable OOXML image part, so the original asset copy was skipped.",
    )
