from __future__ import annotations

import contextlib
import io
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import excel_semantic_md.cli.main as cli_main
from excel_semantic_md.excel import detect_blocks, link_visuals, read_visual_metadata, read_workbook
from excel_semantic_md.render.planner import build_render_plan


FIXTURES = Path(__file__).resolve().parent / "fixtures" / "visuals"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


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
        "source": {
            "drawing_part": "xl/drawings/drawing1.xml",
            "relationship_id": None,
            "target_part": None,
        },
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
        "source": {
            "drawing_part": "xl/drawings/drawing1.xml",
            "relationship_id": None,
            "target_part": None,
        },
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


def test_read_visual_metadata_keeps_sheet_successful_when_drawing_xml_is_malformed(tmp_path: Path) -> None:
    workbook_path = tmp_path / "malformed-drawing.xlsx"
    _copy_workbook_with_replaced_part(
        FIXTURES / "image-visual.xlsx",
        workbook_path,
        "xl/drawings/drawing1.xml",
        b"<xdr:wsDr><xdr:oneCellAnchor>",
    )

    sheet = read_visual_metadata(workbook_path)[0]

    assert sheet.visuals == []
    assert [warning.code for warning in sheet.warnings] == ["drawing_part_parse_failed"]
    assert sheet.warnings[0].details["sheet"] == "Image"
    assert sheet.warnings[0].details["drawing_part"] == "xl/drawings/drawing1.xml"


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


def test_inspect_succeeds_when_drawing_xml_is_malformed(tmp_path: Path) -> None:
    workbook_path = tmp_path / "malformed-drawing.xlsx"
    _copy_workbook_with_replaced_part(
        FIXTURES / "image-visual.xlsx",
        workbook_path,
        "xl/drawings/drawing1.xml",
        b"<xdr:wsDr><xdr:oneCellAnchor>",
    )

    code, stdout, stderr = _run_cli(["inspect", "--input", str(workbook_path)])
    payload = json.loads(stdout)

    assert code == 0
    assert stderr == ""
    assert payload["sheets"][0]["visuals"] == []
    assert payload["sheets"][0]["warnings"][0]["code"] == "drawing_part_parse_failed"


def test_malformed_drawing_warning_does_not_stop_following_sheets(tmp_path: Path) -> None:
    workbook_path = tmp_path / "malformed-drawing-multi-sheet.xlsx"
    _copy_workbook_with_replaced_part_and_empty_second_sheet(
        FIXTURES / "image-visual.xlsx",
        workbook_path,
        "xl/drawings/drawing1.xml",
        b"<xdr:wsDr><xdr:oneCellAnchor>",
    )

    sheets = read_visual_metadata(workbook_path)

    assert [sheet.name for sheet in sheets] == ["Image", "Second"]
    assert [warning.code for warning in sheets[0].warnings] == ["drawing_part_parse_failed"]
    assert sheets[1].visuals == []
    assert sheets[1].warnings == []


def test_read_visual_metadata_warns_when_sheet_drawing_relationships_part_is_missing(tmp_path: Path) -> None:
    workbook_path = tmp_path / "missing-sheet-drawing-rels.xlsx"
    _copy_workbook_without_part(
        FIXTURES / "image-visual.xlsx",
        workbook_path,
        "xl/worksheets/_rels/sheet1.xml.rels",
    )

    sheet = read_visual_metadata(workbook_path)[0]

    assert sheet.visuals == []
    assert [warning.code for warning in sheet.warnings] == ["sheet_drawing_relationships_missing"]


def test_read_visual_metadata_warns_when_chart_target_is_missing(tmp_path: Path) -> None:
    workbook_path = tmp_path / "missing-chart-target.xlsx"
    _copy_workbook_without_part(
        FIXTURES / "chart-visual.xlsx",
        workbook_path,
        "__no_such_part__",
    )
    _rewrite_zip_part_bytes(
        workbook_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda payload: payload.replace(b"/xl/charts/chart1.xml", b"/xl/charts/missing.xml"),
    )

    sheet = read_visual_metadata(workbook_path)[0]

    assert [warning.code for warning in sheet.visuals[0].warnings] == ["chart_part_missing"]


def test_read_visual_metadata_warns_when_chart_part_is_malformed(tmp_path: Path) -> None:
    workbook_path = tmp_path / "malformed-chart.xlsx"
    _copy_workbook_with_replaced_part(
        FIXTURES / "chart-visual.xlsx",
        workbook_path,
        "xl/charts/chart1.xml",
        b"<c:chartSpace>",
    )

    sheet = read_visual_metadata(workbook_path)[0]

    assert [warning.code for warning in sheet.visuals[0].warnings] == ["chart_part_parse_failed"]


def test_build_render_plan_skips_original_image_copy_for_non_image_content_type(tmp_path: Path) -> None:
    workbook_path = tmp_path / "image-non-image-target.xlsx"
    _copy_workbook_with_replaced_content_type(
        FIXTURES / "image-visual.xlsx",
        workbook_path,
        "/xl/media/image1.png",
        "application/octet-stream",
    )

    workbook = link_visuals(detect_blocks(read_workbook(workbook_path)), read_visual_metadata(workbook_path))
    visual_sheet = read_visual_metadata(workbook_path)[0]
    items, warnings, failures = build_render_plan(workbook.sheets[0], visual_sheet)

    assert failures == []
    assert [item.source for item in items] == ["range_copy_picture", "shape_copy_picture"]
    assert [warning.code for warning in warnings] == ["image_original_asset_invalid_content_type"]


def test_build_render_plan_skips_original_image_copy_for_untrusted_image_part(tmp_path: Path) -> None:
    workbook_path = tmp_path / "image-untrusted-target.xlsx"
    _copy_workbook_with_replaced_content_type(
        FIXTURES / "image-visual.xlsx",
        workbook_path,
        "/xl/worksheets/sheet1.xml",
        "image/png",
    )
    _rewrite_zip_part_bytes(
        workbook_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda payload: payload.replace(b"../media/image1.png", b"../worksheets/sheet1.xml"),
    )

    workbook = link_visuals(detect_blocks(read_workbook(workbook_path)), read_visual_metadata(workbook_path))
    visual_sheet = read_visual_metadata(workbook_path)[0]
    items, warnings, failures = build_render_plan(workbook.sheets[0], visual_sheet)

    assert failures == []
    assert [item.source for item in items] == ["range_copy_picture", "shape_copy_picture"]
    assert [warning.code for warning in warnings] == ["image_original_asset_untrusted_part"]


def test_build_render_plan_skips_original_image_copy_when_image_target_is_missing(tmp_path: Path) -> None:
    workbook_path = tmp_path / "image-missing-target.xlsx"
    _copy_workbook_with_rewritten_part(
        FIXTURES / "image-visual.xlsx",
        workbook_path,
        "xl/drawings/_rels/drawing1.xml.rels",
        lambda payload: payload.replace(b'Id="rIdImg1"', b'Id="rIdOther"'),
    )

    workbook = link_visuals(detect_blocks(read_workbook(workbook_path)), read_visual_metadata(workbook_path))
    visual_sheet = read_visual_metadata(workbook_path)[0]
    items, warnings, failures = build_render_plan(workbook.sheets[0], visual_sheet)

    assert failures == []
    assert [warning.code for warning in visual_sheet.visuals[0].warnings] == ["image_target_missing"]
    assert [item.source for item in items] == ["range_copy_picture", "shape_copy_picture"]
    assert [warning.code for warning in warnings] == ["image_original_asset_unavailable"]


def test_build_render_plan_skips_original_image_copy_when_image_part_is_missing(tmp_path: Path) -> None:
    workbook_path = tmp_path / "image-missing-part.xlsx"
    _copy_workbook_without_part(
        FIXTURES / "image-visual.xlsx",
        workbook_path,
        "xl/media/image1.png",
    )

    workbook = link_visuals(detect_blocks(read_workbook(workbook_path)), read_visual_metadata(workbook_path))
    visual_sheet = read_visual_metadata(workbook_path)[0]
    items, warnings, failures = build_render_plan(workbook.sheets[0], visual_sheet)

    assert failures == []
    assert [warning.code for warning in visual_sheet.visuals[0].warnings] == ["image_part_missing"]
    assert [item.source for item in items] == ["range_copy_picture", "shape_copy_picture"]
    assert [warning.code for warning in warnings] == ["image_original_asset_unavailable"]


def test_inspect_includes_visual_metadata_and_linked_visual_blocks() -> None:
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
            "visual_id": None,
            "related_block_id": None,
            "assets": [],
            "warnings": [],
            "text": "Image fixture",
        },
        {
            "id": "s001-b002-image",
            "kind": "image",
            "anchor": {
                "sheet": "Image",
                "start_row": 2,
                "start_col": 4,
                "end_row": 2,
                "end_col": 4,
                "a1": "D2",
            },
            "source": "image",
            "visual_id": "s001-v001-image",
            "related_block_id": "s001-b001-paragraph",
            "assets": [],
            "warnings": [],
            "alt_text": "Company logo",
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


def _copy_workbook_with_replaced_part(source: Path, target: Path, part_name: str, payload: bytes) -> None:
    with zipfile.ZipFile(source) as source_archive, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as target_archive:
        for item in source_archive.infolist():
            target_archive.writestr(item, payload if item.filename == part_name else source_archive.read(item.filename))


def _copy_workbook_with_rewritten_part(source: Path, target: Path, part_name: str, rewriter) -> None:
    with zipfile.ZipFile(source) as source_archive, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as target_archive:
        for item in source_archive.infolist():
            payload = source_archive.read(item.filename)
            target_archive.writestr(item, rewriter(payload) if item.filename == part_name else payload)


def _copy_workbook_without_part(source: Path, target: Path, part_name: str) -> None:
    with zipfile.ZipFile(source) as source_archive, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as target_archive:
        for item in source_archive.infolist():
            if item.filename == part_name:
                continue
            target_archive.writestr(item, source_archive.read(item.filename))


def _copy_workbook_with_replaced_content_type(source: Path, target: Path, part_name: str, content_type: str) -> None:
    with zipfile.ZipFile(source) as source_archive, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as target_archive:
        for item in source_archive.infolist():
            payload = source_archive.read(item.filename)
            if item.filename == "[Content_Types].xml":
                payload = _rewrite_content_type(payload, part_name, content_type)
            target_archive.writestr(item, payload)


def _rewrite_zip_part_bytes(path: Path, part_name: str, rewriter) -> None:
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as source_archive, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target_archive:
        for item in source_archive.infolist():
            payload = source_archive.read(item.filename)
            if item.filename == part_name:
                payload = rewriter(payload)
            target_archive.writestr(item, payload)
    temp_path.replace(path)


def _rewrite_content_type(payload: bytes, part_name: str, content_type: str) -> bytes:
    root = ElementTree.fromstring(payload)
    found = False
    for node in root:
        if node.attrib.get("PartName") == part_name:
            node.attrib["ContentType"] = content_type
            found = True
            break
    if not found:
        ElementTree.SubElement(
            root,
            "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
            {"PartName": part_name, "ContentType": content_type},
        )
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)


def _copy_workbook_with_replaced_part_and_empty_second_sheet(source: Path, target: Path, part_name: str, payload: bytes) -> None:
    with zipfile.ZipFile(source) as source_archive, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as target_archive:
        for item in source_archive.infolist():
            data = source_archive.read(item.filename)
            if item.filename == part_name:
                data = payload
            elif item.filename == "xl/workbook.xml":
                data = _workbook_with_second_sheet(data)
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = _workbook_relationships_with_second_sheet(data)
            elif item.filename == "[Content_Types].xml":
                data = _content_types_with_second_sheet(data)
            target_archive.writestr(item, data)
        target_archive.writestr(
            "xl/worksheets/sheet2.xml",
            (
                f'<worksheet xmlns="{MAIN_NS}">'
                "<sheetData>"
                '<row r="1"><c r="A1" t="inlineStr"><is><t>Second sheet</t></is></c></row>'
                "</sheetData>"
                "</worksheet>"
            ),
        )


def _workbook_with_second_sheet(payload: bytes) -> bytes:
    root = ElementTree.fromstring(payload)
    sheets_node = root.find(f"{{{MAIN_NS}}}sheets")
    assert sheets_node is not None
    ElementTree.SubElement(
        sheets_node,
        f"{{{MAIN_NS}}}sheet",
        {"name": "Second", "sheetId": "2", f"{{{REL_NS}}}id": "rIdSecond"},
    )
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)


def _workbook_relationships_with_second_sheet(payload: bytes) -> bytes:
    root = ElementTree.fromstring(payload)
    ElementTree.SubElement(
        root,
        f"{{{PKG_REL_NS}}}Relationship",
        {
            "Id": "rIdSecond",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "Target": "worksheets/sheet2.xml",
        },
    )
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)


def _content_types_with_second_sheet(payload: bytes) -> bytes:
    root = ElementTree.fromstring(payload)
    ElementTree.SubElement(
        root,
        f"{{{CONTENT_TYPES_NS}}}Override",
        {
            "PartName": "/xl/worksheets/sheet2.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        },
    )
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)
