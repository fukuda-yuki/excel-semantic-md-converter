from __future__ import annotations

from pathlib import Path

from excel_semantic_md.excel import detect_blocks, link_visuals, read_visual_metadata, read_workbook
from excel_semantic_md.excel.ooxml_visual_reader import (
    AssetCandidate,
    SheetVisualResult,
    VisualAnchor,
    VisualAnchorPoint,
    VisualElement,
    VisualSource,
)
from excel_semantic_md.models import HeadingBlock, ParagraphBlock, Rect, SheetModel, SourceKind, WorkbookModel


FIXTURES = Path(__file__).resolve().parent / "fixtures" / "visuals"


def test_links_adjacent_shape_to_table_fixture() -> None:
    linked = _linked_workbook("table-shape-visual.xlsx")
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["kind"] for block in blocks] == ["table", "shape"]
    assert blocks[1]["visual_id"] == "s001-v001-shape"
    assert blocks[1]["related_block_id"] == "s001-b001-table"
    assert blocks[1]["text"] == "Quarterly callout"


def test_links_adjacent_image_to_table_fixture() -> None:
    linked = _linked_workbook("table-image-visual.xlsx")
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["kind"] for block in blocks] == ["table", "image"]
    assert blocks[1]["visual_id"] == "s001-v001-image"
    assert blocks[1]["related_block_id"] == "s001-b001-table"
    assert blocks[1]["alt_text"] == "Company logo"


def test_links_adjacent_chart_to_table_fixture() -> None:
    linked = _linked_workbook("chart-visual.xlsx")
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["kind"] for block in blocks] == ["table", "chart"]
    assert blocks[1]["visual_id"] == "s001-v001-chart"
    assert blocks[1]["related_block_id"] == "s001-b001-table"
    assert blocks[1]["title"] == "Quarterly Sales"


def test_links_visual_to_heading_scope_until_next_heading() -> None:
    block_model = WorkbookModel(
        sheets=[
            SheetModel(
                sheet_index=1,
                name="Scope",
                blocks=[
                    HeadingBlock(
                        id="pending",
                        anchor=Rect(sheet="Scope", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                        source=SourceKind.CELLS,
                        text="Overview",
                        level=1,
                    ),
                    ParagraphBlock(
                        id="pending",
                        anchor=Rect(sheet="Scope", start_row=8, start_col=1, end_row=8, end_col=1, a1="A8"),
                        source=SourceKind.CELLS,
                        text="Later body",
                    ),
                    HeadingBlock(
                        id="pending",
                        anchor=Rect(sheet="Scope", start_row=10, start_col=1, end_row=10, end_col=1, a1="A10"),
                        source=SourceKind.CELLS,
                        text="Next",
                        level=1,
                    ),
                ],
            )
        ]
    )
    visual_results = [
        SheetVisualResult(
            sheet_index=1,
            name="Scope",
            visuals=[
                VisualElement(
                    id="s001-v001-image",
                    kind="image",
                    anchor=VisualAnchor(
                        anchor_type="oneCellAnchor",
                        from_point=VisualAnchorPoint(row=2, col=4),
                        a1="D2",
                    ),
                    source=VisualSource(drawing_part="xl/drawings/drawing1.xml"),
                    asset_candidate=AssetCandidate(kind="image"),
                    alt_text="Scoped visual",
                )
            ],
        )
    ]

    linked = link_visuals(block_model, visual_results)
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["kind"] for block in blocks] == ["heading", "image", "paragraph", "heading"]
    assert blocks[1]["related_block_id"] == "s001-b001-heading"
    assert blocks[1]["visual_id"] == "s001-v001-image"


def test_links_visual_to_heading_scope_after_intermediate_block() -> None:
    block_model = WorkbookModel(
        sheets=[
            SheetModel(
                sheet_index=1,
                name="ScopeLate",
                blocks=[
                    HeadingBlock(
                        id="pending",
                        anchor=Rect(sheet="ScopeLate", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                        source=SourceKind.CELLS,
                        text="Overview",
                        level=1,
                    ),
                    ParagraphBlock(
                        id="pending",
                        anchor=Rect(sheet="ScopeLate", start_row=4, start_col=1, end_row=4, end_col=1, a1="A4"),
                        source=SourceKind.CELLS,
                        text="Body block",
                    ),
                    HeadingBlock(
                        id="pending",
                        anchor=Rect(sheet="ScopeLate", start_row=10, start_col=1, end_row=10, end_col=1, a1="A10"),
                        source=SourceKind.CELLS,
                        text="Next",
                        level=1,
                    ),
                ],
            )
        ]
    )
    visual_results = [
        SheetVisualResult(
            sheet_index=1,
            name="ScopeLate",
            visuals=[
                VisualElement(
                    id="s001-v001-image",
                    kind="image",
                    anchor=VisualAnchor(
                        anchor_type="oneCellAnchor",
                        from_point=VisualAnchorPoint(row=8, col=4),
                        a1="D8",
                    ),
                    source=VisualSource(drawing_part="xl/drawings/drawing1.xml"),
                    asset_candidate=AssetCandidate(kind="image"),
                    alt_text="Late scoped visual",
                )
            ],
        )
    ]

    linked = link_visuals(block_model, visual_results)
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["kind"] for block in blocks] == ["heading", "paragraph", "image", "heading"]
    assert blocks[2]["related_block_id"] == "s001-b001-heading"


def test_assigns_synthetic_anchor_and_warning_for_absolute_anchor_visual() -> None:
    block_model = WorkbookModel(
        sheets=[
            SheetModel(
                sheet_index=1,
                name="Absolute",
                blocks=[
                    ParagraphBlock(
                        id="pending",
                        anchor=Rect(sheet="Absolute", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                        source=SourceKind.CELLS,
                        text="Anchor me",
                    )
                ],
            )
        ]
    )
    visual_results = [
        SheetVisualResult(
            sheet_index=1,
            name="Absolute",
            visuals=[
                VisualElement(
                    id="s001-v001-image",
                    kind="image",
                    anchor=VisualAnchor(anchor_type="absoluteAnchor"),
                    source=VisualSource(drawing_part="xl/drawings/drawing1.xml"),
                    asset_candidate=AssetCandidate(kind="image"),
                    alt_text="Absolute visual",
                )
            ],
        )
    ]

    linked = link_visuals(block_model, visual_results)
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["kind"] for block in blocks] == ["paragraph", "image"]
    assert blocks[1]["anchor"]["a1"] == "A2"
    assert blocks[1]["related_block_id"] is None
    assert [warning["code"] for warning in blocks[1]["warnings"]] == ["visual_anchor_not_cell_addressable"]


def test_places_synthetic_anchor_after_lower_real_visuals() -> None:
    block_model = WorkbookModel(
        sheets=[
            SheetModel(
                sheet_index=1,
                name="SyntheticOrder",
                blocks=[
                    ParagraphBlock(
                        id="pending",
                        anchor=Rect(sheet="SyntheticOrder", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                        source=SourceKind.CELLS,
                        text="Top",
                    )
                ],
            )
        ]
    )
    visual_results = [
        SheetVisualResult(
            sheet_index=1,
            name="SyntheticOrder",
            visuals=[
                VisualElement(
                    id="s001-v001-image",
                    kind="image",
                    anchor=VisualAnchor(
                        anchor_type="oneCellAnchor",
                        from_point=VisualAnchorPoint(row=20, col=4),
                        a1="D20",
                    ),
                    source=VisualSource(drawing_part="xl/drawings/drawing1.xml"),
                    asset_candidate=AssetCandidate(kind="image"),
                    alt_text="Lower real visual",
                ),
                VisualElement(
                    id="s001-v002-image",
                    kind="image",
                    anchor=VisualAnchor(anchor_type="absoluteAnchor"),
                    source=VisualSource(drawing_part="xl/drawings/drawing1.xml"),
                    asset_candidate=AssetCandidate(kind="image"),
                    alt_text="Synthetic visual",
                ),
            ],
        )
    ]

    linked = link_visuals(block_model, visual_results)
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["kind"] for block in blocks] == ["paragraph", "image", "image"]
    assert blocks[1]["anchor"]["a1"] == "D20"
    assert blocks[2]["anchor"]["a1"] == "A21"
    assert blocks[2]["related_block_id"] is None


def test_chooses_nearest_block_when_no_adjacent_or_heading_scope_match() -> None:
    block_model = WorkbookModel(
        sheets=[
            SheetModel(
                sheet_index=1,
                name="Nearest",
                blocks=[
                    ParagraphBlock(
                        id="pending",
                        anchor=Rect(sheet="Nearest", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                        source=SourceKind.CELLS,
                        text="Left",
                    ),
                    ParagraphBlock(
                        id="pending",
                        anchor=Rect(sheet="Nearest", start_row=10, start_col=10, end_row=10, end_col=10, a1="J10"),
                        source=SourceKind.CELLS,
                        text="Right",
                    ),
                ],
            )
        ]
    )
    visual_results = [
        SheetVisualResult(
            sheet_index=1,
            name="Nearest",
            visuals=[
                VisualElement(
                    id="s001-v001-image",
                    kind="image",
                    anchor=VisualAnchor(
                        anchor_type="oneCellAnchor",
                        from_point=VisualAnchorPoint(row=3, col=3),
                        a1="C3",
                    ),
                    source=VisualSource(drawing_part="xl/drawings/drawing1.xml"),
                    asset_candidate=AssetCandidate(kind="image"),
                    alt_text="Nearest visual",
                )
            ],
        )
    ]

    linked = link_visuals(block_model, visual_results)
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["kind"] for block in blocks] == ["paragraph", "image", "paragraph"]
    assert blocks[1]["related_block_id"] == "s001-b001-paragraph"


def test_ignores_unknown_visual_when_building_blocks() -> None:
    linked = _linked_workbook("table-shape-visual.xlsx")
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert [block["visual_id"] for block in blocks if block["visual_id"] is not None] == ["s001-v001-shape"]


def test_link_visuals_does_not_mutate_input_block_model() -> None:
    block_model = WorkbookModel(
        sheets=[
            SheetModel(
                sheet_index=1,
                name="Immutable",
                blocks=[
                    ParagraphBlock(
                        id="pending",
                        anchor=Rect(sheet="Immutable", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
                        source=SourceKind.CELLS,
                        text="Original",
                    )
                ],
            )
        ]
    )
    original = block_model.to_dict()
    visual_results = [
        SheetVisualResult(
            sheet_index=1,
            name="Immutable",
            visuals=[
                VisualElement(
                    id="s001-v001-image",
                    kind="image",
                    anchor=VisualAnchor(
                        anchor_type="oneCellAnchor",
                        from_point=VisualAnchorPoint(row=2, col=4),
                        a1="D2",
                    ),
                    source=VisualSource(drawing_part="xl/drawings/drawing1.xml"),
                    asset_candidate=AssetCandidate(kind="image"),
                    alt_text="Linked visual",
                )
            ],
        )
    ]

    _linked = link_visuals(block_model, visual_results)

    assert block_model.to_dict() == original


def test_normalizes_visual_block_a1_from_numeric_anchor_bounds() -> None:
    block_model = WorkbookModel(sheets=[SheetModel(sheet_index=1, name="Normalize", blocks=[])])
    visual_results = [
        SheetVisualResult(
            sheet_index=1,
            name="Normalize",
            visuals=[
                VisualElement(
                    id="s001-v001-image",
                    kind="image",
                    anchor=VisualAnchor(
                        anchor_type="twoCellAnchor",
                        from_point=VisualAnchorPoint(row=3, col=3),
                        to_point=VisualAnchorPoint(row=5, col=4),
                        a1="z99",
                    ),
                    source=VisualSource(drawing_part="xl/drawings/drawing1.xml"),
                    asset_candidate=AssetCandidate(kind="image"),
                    alt_text="Normalized visual",
                )
            ],
        )
    ]

    linked = link_visuals(block_model, visual_results)
    blocks = [block.to_dict() for block in linked.sheets[0].blocks]

    assert blocks[0]["anchor"]["a1"] == "C3:D5"


def _linked_workbook(name: str) -> WorkbookModel:
    path = FIXTURES / name
    return link_visuals(detect_blocks(read_workbook(path)), read_visual_metadata(path))
