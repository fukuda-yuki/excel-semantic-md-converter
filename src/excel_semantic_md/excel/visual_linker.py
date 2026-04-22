"""Link OOXML visuals to detected blocks for Phase 1 inspect output."""

from __future__ import annotations

from dataclasses import dataclass

from excel_semantic_md.excel.ooxml_visual_reader import SheetVisualResult, VisualElement
from excel_semantic_md.models import (
    Block,
    ChartBlock,
    HeadingBlock,
    ImageBlock,
    Rect,
    ShapeBlock,
    SheetModel,
    SourceKind,
    WarningInfo,
    WorkbookModel,
    make_block_id,
)


@dataclass(frozen=True)
class _HeadingScope:
    heading: HeadingBlock
    start_row_exclusive: int
    end_row_exclusive: int | None = None


_SOURCE_ORDER = {
    SourceKind.CELLS: 0,
    SourceKind.SHAPE: 1,
    SourceKind.IMAGE: 2,
    SourceKind.CHART: 3,
    SourceKind.UNKNOWN: 4,
}


def link_visuals(block_model: WorkbookModel, visual_results: list[SheetVisualResult]) -> WorkbookModel:
    """Return a workbook model whose sheet blocks include linked visual-origin blocks."""

    visual_results_by_index = {result.sheet_index: result for result in visual_results}
    linked_sheets: list[SheetModel] = []
    for sheet in block_model.sheets:
        linked_sheets.append(_link_sheet_visuals(sheet, visual_results_by_index.get(sheet.sheet_index)))
    return WorkbookModel(
        sheets=linked_sheets,
        input_file_name=block_model.input_file_name,
        schema_version=block_model.schema_version,
    )


def _link_sheet_visuals(sheet: SheetModel, visual_result: SheetVisualResult | None) -> SheetModel:
    cloned_blocks = [_clone_block(block) for block in sheet.blocks]
    if visual_result is None:
        return SheetModel(
            sheet_index=sheet.sheet_index,
            name=sheet.name,
            blocks=cloned_blocks,
            failures=list(sheet.failures),
            warnings=list(sheet.warnings),
        )

    combined_blocks = list(cloned_blocks)
    cell_blocks = [block for block in cloned_blocks if block.source == SourceKind.CELLS]
    heading_scopes = _heading_scopes(cell_blocks)
    max_addressable_row = max(
        [block.anchor.end_row for block in cloned_blocks]
        + [
            rect.end_row
            for rect in (_visual_anchor_rect(visual, sheet.name) for visual in visual_result.visuals)
            if rect is not None
        ],
        default=0,
    )
    next_fallback_row = max_addressable_row + 1
    related_targets_by_visual_block: dict[int, int] = {}

    for visual in visual_result.visuals:
        if visual.kind not in {"shape", "image", "chart"}:
            continue

        anchor_rect = _visual_anchor_rect(visual, sheet.name)
        rect_warnings: list[WarningInfo] = []
        anchor_is_synthetic = False
        if anchor_rect is None:
            anchor_rect = _synthetic_rect(sheet.name, next_fallback_row)
            next_fallback_row += 1
            anchor_is_synthetic = True
            rect_warnings.append(
                WarningInfo(
                    code="visual_anchor_not_cell_addressable",
                    message="Visual anchor could not be converted to a cell range, so a synthetic standalone anchor was assigned.",
                    details={"visual_id": visual.id, "anchor_type": visual.anchor.anchor_type},
                )
            )

        visual_block = _block_from_visual(visual, anchor_rect, rect_warnings)
        if not anchor_is_synthetic:
            related = _resolve_related_block(anchor_rect, cell_blocks, heading_scopes)
            if related is not None:
                related_targets_by_visual_block[id(visual_block)] = id(related)
        combined_blocks.append(visual_block)

    sorted_blocks = sorted(combined_blocks, key=_sort_key)
    for block_index, block in enumerate(sorted_blocks, start=1):
        block.id = make_block_id(sheet.sheet_index, block_index, block.kind)

    block_ids_by_object = {id(block): block.id for block in sorted_blocks}
    for block in sorted_blocks:
        block.related_block_id = None
        related_target = related_targets_by_visual_block.get(id(block))
        if related_target is not None:
            block.related_block_id = block_ids_by_object[related_target]

    return SheetModel(
        sheet_index=sheet.sheet_index,
        name=sheet.name,
        blocks=sorted_blocks,
        failures=list(sheet.failures),
        warnings=list(sheet.warnings),
    )


def _heading_scopes(blocks: list[Block]) -> list[_HeadingScope]:
    ordered = sorted(blocks, key=_sort_key)
    heading_positions = [
        (index, block)
        for index, block in enumerate(ordered)
        if isinstance(block, HeadingBlock)
    ]
    scopes: list[_HeadingScope] = []
    for offset, (_index, block) in enumerate(heading_positions):
        next_row = heading_positions[offset + 1][1].anchor.start_row if offset + 1 < len(heading_positions) else None
        scopes.append(
            _HeadingScope(
                heading=block,
                start_row_exclusive=block.anchor.end_row,
                end_row_exclusive=next_row,
            )
        )
    return scopes


def _resolve_related_block(
    anchor_rect: Rect,
    cell_blocks: list[Block],
    heading_scopes: list[_HeadingScope],
) -> Block | None:
    if not cell_blocks:
        return None

    ordered_blocks = sorted(cell_blocks, key=_sort_key)
    adjacent = [block for block in ordered_blocks if _rect_distance(anchor_rect, block.anchor) <= 1]
    if adjacent:
        return min(adjacent, key=lambda block: (_rect_distance(anchor_rect, block.anchor), _sort_key(block)))

    heading_match = _heading_scope_match(anchor_rect, heading_scopes)
    if heading_match is not None:
        return heading_match

    return min(ordered_blocks, key=lambda block: (_rect_distance(anchor_rect, block.anchor), _sort_key(block)))


def _heading_scope_match(
    anchor_rect: Rect,
    heading_scopes: list[_HeadingScope],
) -> HeadingBlock | None:
    for scope in heading_scopes:
        if anchor_rect.start_row <= scope.start_row_exclusive:
            continue
        if scope.end_row_exclusive is not None and anchor_rect.start_row >= scope.end_row_exclusive:
            continue
        return scope.heading
    return None


def _visual_anchor_rect(visual: VisualElement, sheet_name: str) -> Rect | None:
    from_point = visual.anchor.from_point
    to_point = visual.anchor.to_point
    if from_point is None or from_point.row is None or from_point.col is None:
        return None

    if visual.anchor.anchor_type == "oneCellAnchor":
        start_row = end_row = from_point.row
        start_col = end_col = from_point.col
    elif visual.anchor.anchor_type == "twoCellAnchor":
        if to_point is None or to_point.row is None or to_point.col is None:
            return None
        start_row = min(from_point.row, to_point.row)
        end_row = max(from_point.row, to_point.row)
        start_col = min(from_point.col, to_point.col)
        end_col = max(from_point.col, to_point.col)
    else:
        return None

    return Rect(
        sheet=sheet_name,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        a1=_rect_a1(start_row, start_col, end_row, end_col),
    )


def _synthetic_rect(sheet_name: str, row: int) -> Rect:
    return Rect(sheet=sheet_name, start_row=row, start_col=1, end_row=row, end_col=1, a1=f"A{row}")


def _block_from_visual(
    visual: VisualElement,
    anchor: Rect,
    rect_warnings: list[WarningInfo],
) -> Block:
    warnings = [_warning_from_read_warning(item) for item in visual.warnings]
    warnings.extend(rect_warnings)
    if visual.kind == "shape":
        return ShapeBlock(
            id="pending",
            anchor=anchor,
            source=SourceKind.SHAPE,
            visual_id=visual.id,
            text=visual.text or "",
            shape_type=visual.shape_type,
            warnings=warnings,
        )
    if visual.kind == "image":
        return ImageBlock(
            id="pending",
            anchor=anchor,
            source=SourceKind.IMAGE,
            visual_id=visual.id,
            alt_text=visual.alt_text,
            warnings=warnings,
        )
    return ChartBlock(
        id="pending",
        anchor=anchor,
        source=SourceKind.CHART,
        visual_id=visual.id,
        title=visual.title,
        metadata={"series": [series.to_dict() for series in visual.series]},
        warnings=warnings,
    )


def _warning_from_read_warning(warning: object) -> WarningInfo:
    code = getattr(warning, "code")
    message = getattr(warning, "message")
    details = dict(getattr(warning, "details", {}))
    return WarningInfo(code=code, message=message, details=details)


def _clone_block(block: Block) -> Block:
    return Block.from_dict(block.to_dict())


def _sort_key(block: Block) -> tuple[int, int, int, int, int, str]:
    return (*_rect_sort_key(block.anchor), _SOURCE_ORDER[block.source], block.kind.value)


def _rect_sort_key(anchor: Rect) -> tuple[int, int, int, int]:
    return (anchor.start_row, anchor.start_col, anchor.end_row, anchor.end_col)


def _rect_distance(left: Rect, right: Rect) -> int:
    row_gap = max(0, left.start_row - right.end_row - 1, right.start_row - left.end_row - 1)
    col_gap = max(0, left.start_col - right.end_col - 1, right.start_col - left.end_col - 1)
    return max(row_gap, col_gap)


def _rect_a1(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
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


__all__ = ["link_visuals"]
