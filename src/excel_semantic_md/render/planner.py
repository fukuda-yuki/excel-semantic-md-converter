"""Build render plans for Excel COM confirmation."""

from __future__ import annotations

from excel_semantic_md.excel.ooxml_visual_reader import SheetVisualResult, VisualElement
from excel_semantic_md.models import AssetRole, Block, FailureInfo, SheetModel, SourceKind, WarningInfo
from excel_semantic_md.render.types import RenderPlanItem


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
                target_part = _image_target_part(visual)
                if target_part is None:
                    warnings.append(
                        WarningInfo(
                            code="image_original_asset_unavailable",
                            message="The linked image does not expose an OOXML target part, so the original asset copy was skipped.",
                            details={"block_id": block.id, "visual_id": block.visual_id},
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


def _image_target_part(visual: VisualElement | None) -> str | None:
    if visual is None or visual.kind != "image":
        return None
    return visual.source.target_part
