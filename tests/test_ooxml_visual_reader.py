from __future__ import annotations

import contextlib
import io
import json
from pathlib import Path

import excel_semantic_md.cli.main as cli_main
from excel_semantic_md.excel import read_visual_metadata


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


def test_read_visual_metadata_returns_empty_visuals_for_sheet_without_drawings() -> None:
    sheet = read_visual_metadata(FIXTURES / "no-visuals.xlsx")[0]

    assert sheet.name == "Plain"
    assert sheet.visuals == []
    assert sheet.warnings == []


def test_read_visual_metadata_reads_embedded_image_and_alt_text() -> None:
    sheet = read_visual_metadata(FIXTURES / "image-visual.xlsx")[0]
    visual = sheet.visuals[0].to_dict()

    assert sheet.warnings == []
    assert visual == {
        "id": "s001-v001-image",
        "kind": "image",
        "anchor": {
            "anchor_type": "oneCellAnchor",
            "from": {
                "row": 2,
                "col": 4,
                "row_offset_emu": 240000,
                "col_offset_emu": 120000,
            },
            "a1": "D2",
        },
        "source": {
            "drawing_part": "xl/drawings/drawing1.xml",
            "relationship_id": "rIdImg1",
            "target_part": "xl/media/image1.png",
        },
        "asset_candidate": {
            "kind": "image",
            "source_part": "xl/media/image1.png",
            "extension": ".png",
            "content_type": "image/png",
        },
        "warnings": [],
        "alt_text": "Company logo",
    }


def test_read_visual_metadata_reads_chart_metadata_and_series_cache() -> None:
    sheet = read_visual_metadata(FIXTURES / "chart-visual.xlsx")[0]
    visual = sheet.visuals[0].to_dict()

    assert sheet.warnings == []
    assert visual["kind"] == "chart"
    assert visual["title"] == "Quarterly Sales"
    assert visual["anchor"]["anchor_type"] == "oneCellAnchor"
    assert visual["anchor"]["a1"] == "D2"
    assert visual["source"]["target_part"] == "xl/charts/chart1.xml"
    assert visual["asset_candidate"] == {
        "kind": "chart",
        "source_part": "xl/charts/chart1.xml",
        "extension": ".xml",
        "content_type": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    }
    assert visual["series"] == [
        {
            "name": "'Chart'!B1",
            "categories": ["North", "South"],
            "values": ["12", "18"],
            "category_ref": "'Chart'!$A$2:$A$3",
            "value_ref": "'Chart'!$B$2:$B$3",
        }
    ]


def test_read_visual_metadata_reads_shape_and_unknown_group_shape() -> None:
    sheet = read_visual_metadata(FIXTURES / "shape-unknown.xlsx")[0]
    visuals = [item.to_dict() for item in sheet.visuals]

    assert sheet.warnings == []
    assert visuals[0] == {
        "id": "s001-v001-shape",
        "kind": "shape",
        "anchor": {
            "anchor_type": "twoCellAnchor",
            "from": {"row": 2, "col": 2},
            "to": {"row": 4, "col": 4},
            "a1": "B2:D4",
        },
        "source": {"drawing_part": "xl/drawings/drawing1.xml"},
        "asset_candidate": {
            "kind": "shape",
            "source_part": "xl/drawings/drawing1.xml",
            "extension": ".xml",
            "content_type": "application/vnd.openxmlformats-officedocument.drawing+xml",
        },
        "warnings": [],
        "text": "Quarterly callout",
        "shape_type": "rect",
    }
    assert visuals[1] == {
        "id": "s001-v002-unknown",
        "kind": "unknown",
        "anchor": {
            "anchor_type": "twoCellAnchor",
            "from": {"row": 6, "col": 1},
            "to": {"row": 8, "col": 3},
            "a1": "A6:C8",
        },
        "source": {"drawing_part": "xl/drawings/drawing1.xml"},
        "asset_candidate": {
            "kind": "unknown",
            "source_part": "xl/drawings/drawing1.xml",
            "extension": ".xml",
            "content_type": "application/vnd.openxmlformats-officedocument.drawing+xml",
        },
        "warnings": [
            {
                "code": "unsupported_visual_element",
                "message": "Group shape is not fully interpreted in Phase 1.",
                "details": {
                    "drawing_part": "xl/drawings/drawing1.xml",
                    "unknown_kind": "group_shape",
                },
            }
        ],
        "unknown_kind": "group_shape",
        "description": "Group shape is not fully interpreted in Phase 1.",
    }


def test_read_visual_metadata_keeps_sheet_successful_when_drawing_target_is_missing() -> None:
    sheet = read_visual_metadata(FIXTURES / "broken-drawing-rel.xlsx")[0]

    assert sheet.visuals == []
    assert [warning.to_dict() for warning in sheet.warnings] == [
        {
            "code": "drawing_part_missing",
            "message": "Drawing part referenced by the sheet was not found.",
            "details": {"sheet": "Broken", "drawing_part": "xl/drawings/missing-drawing.xml"},
        }
    ]


def test_read_visual_metadata_supports_xlsm_without_executing_macros() -> None:
    sheet = read_visual_metadata(FIXTURES / "image-visual.xlsm")[0]
    visual = sheet.visuals[0].to_dict()

    assert sheet.warnings == []
    assert visual["kind"] == "image"
    assert visual["source"]["target_part"] == "xl/media/image1.png"
    assert visual["alt_text"] == "Company logo"


def test_inspect_includes_visuals_and_visual_warnings() -> None:
    code, stdout, stderr = _run_cli(["inspect", "--input", str(FIXTURES / "broken-drawing-rel.xlsx")])

    payload = json.loads(stdout)
    assert code == 0
    assert stderr == ""
    assert payload["sheets"][0]["visuals"] == []
    assert payload["sheets"][0]["warnings"] == [
        {
            "code": "drawing_part_missing",
            "message": "Drawing part referenced by the sheet was not found.",
            "details": {"sheet": "Broken", "drawing_part": "xl/drawings/missing-drawing.xml"},
        }
    ]


def test_inspect_includes_visual_metadata_without_changing_existing_blocks() -> None:
    code, stdout, stderr = _run_cli(["inspect", "--input", str(FIXTURES / "image-visual.xlsx")])

    payload = json.loads(stdout)
    assert code == 0
    assert stderr == ""
    assert payload["sheets"][0]["cells"] == [{"row": 1, "col": 1, "a1": "A1", "text": "Image fixture"}]
    assert payload["sheets"][0]["blocks"] == [
        {
            "id": "s001-b001-paragraph",
            "kind": "paragraph",
            "anchor": {
                "sheet": "Image",
                "start_row": 1,
                "start_col": 1,
                "end_row": 1,
                "end_col": 1,
                "a1": "A1",
            },
            "source": "cells",
            "assets": [],
            "warnings": [],
            "text": "Image fixture",
        }
    ]
    assert payload["sheets"][0]["visuals"] == [
        {
            "id": "s001-v001-image",
            "kind": "image",
            "anchor": {
                "anchor_type": "oneCellAnchor",
                "from": {
                    "row": 2,
                    "col": 4,
                    "row_offset_emu": 240000,
                    "col_offset_emu": 120000,
                },
                "a1": "D2",
            },
            "source": {
                "drawing_part": "xl/drawings/drawing1.xml",
                "relationship_id": "rIdImg1",
                "target_part": "xl/media/image1.png",
            },
            "asset_candidate": {
                "kind": "image",
                "source_part": "xl/media/image1.png",
                "extension": ".png",
                "content_type": "image/png",
            },
            "warnings": [],
            "alt_text": "Company logo",
        }
    ]
