import importlib
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from excel_semantic_md.models import (  # noqa: E402
    AssetKind,
    AssetRef,
    AssetRole,
    Block,
    BlockKind,
    FailureInfo,
    Rect,
    ShapeBlock,
    SheetModel,
    SourceKind,
    TableBlock,
    UnknownBlock,
    UnknownInfo,
    WorkbookModel,
    make_asset_path,
    make_block_id,
)


def test_rect_validates_1_based_bounds_and_order():
    rect = Rect(sheet="Sheet1", start_row=1, start_col=2, end_row=3, end_col=4, a1="B1:D3")

    assert rect.to_dict() == {
        "sheet": "Sheet1",
        "start_row": 1,
        "start_col": 2,
        "end_row": 3,
        "end_col": 4,
        "a1": "B1:D3",
    }

    with pytest.raises(ValueError, match="start_row"):
        Rect(sheet="Sheet1", start_row=0, start_col=1, end_row=1, end_col=1, a1="A0:A1")
    with pytest.raises(ValueError, match="start_col"):
        Rect(sheet="Sheet1", start_row=1, start_col=0, end_row=1, end_col=1, a1="A1")
    with pytest.raises(ValueError, match="start_row"):
        Rect(sheet="Sheet1", start_row=2, start_col=1, end_row=1, end_col=1, a1="A2:A1")
    with pytest.raises(ValueError, match="start_col"):
        Rect(sheet="Sheet1", start_row=1, start_col=2, end_row=1, end_col=1, a1="B1:A1")
    with pytest.raises(ValueError, match="sheet"):
        Rect(sheet="", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1")
    with pytest.raises(ValueError, match="a1"):
        Rect(sheet="Sheet1", start_row=1, start_col=1, end_row=1, end_col=1, a1="")


def test_table_block_round_trip_uses_snake_case_and_enum_values():
    block = TableBlock(
        id="s001-b001-table",
        anchor=Rect(sheet="Sheet1", start_row=1, start_col=1, end_row=2, end_col=2, a1="A1:B2"),
        source=SourceKind.CELLS,
        rows=[["Name", "Value"], ["A", "1"]],
        header_rows=1,
        header_cols=0,
    )

    data = block.to_dict()

    assert data["kind"] == "table"
    assert data["source"] == "cells"
    assert data["visual_id"] is None
    assert data["related_block_id"] is None
    assert "header_rows" in data
    assert "headerRows" not in json.dumps(data)
    assert Block.from_dict(data) == block


def test_shape_block_round_trip_with_asset_ref():
    block = ShapeBlock(
        id="s001-b002-shape",
        anchor=Rect(sheet="Sheet1", start_row=4, start_col=1, end_row=6, end_col=3, a1="A4:C6"),
        source=SourceKind.SHAPE,
        visual_id="s001-v001-shape",
        related_block_id="s001-b001-table",
        text="Important note",
        shape_type="text_box",
        assets=[
            AssetRef(
                path="assets/sheet-001/s001-b002-shape-001.png",
                kind=AssetKind.SHAPE,
                role=AssetRole.MARKDOWN,
                description="Rendered shape",
            )
        ],
    )

    data = block.to_dict()

    assert data["assets"][0]["kind"] == "shape"
    assert data["assets"][0]["role"] == "markdown"
    assert data["visual_id"] == "s001-v001-shape"
    assert data["related_block_id"] == "s001-b001-table"
    assert Block.from_dict(data) == block


def test_unknown_block_and_workbook_round_trip_include_sheet_failures():
    unknown = UnknownBlock(
        id="s002-b001-unknown",
        anchor=Rect(sheet="Second", start_row=1, start_col=1, end_row=1, end_col=1, a1="A1"),
        source=SourceKind.UNKNOWN,
        unknown=UnknownInfo(
            kind="ole",
            description="Unsupported embedded object",
            details={"relationship_id": "rId7"},
        ),
    )
    workbook = WorkbookModel(
        input_file_name="book.xlsx",
        sheets=[
            SheetModel(
                sheet_index=2,
                name="Second",
                blocks=[unknown],
                failures=[FailureInfo(stage="render", message="Chart export failed")],
            )
        ],
    )

    data = workbook.to_dict()

    assert data["schema_version"] == "phase1.0"
    assert data["sheets"][0]["failures"][0]["stage"] == "render"
    assert WorkbookModel.from_dict(data) == workbook


def test_block_id_and_asset_path_are_stable_and_1_based():
    block_id = make_block_id(sheet_index=1, block_index=2, kind=BlockKind.SHAPE)

    assert block_id == "s001-b002-shape"
    assert make_asset_path(sheet_index=1, block_id=block_id, asset_kind=AssetKind.SHAPE, asset_index=3) == (
        "assets/sheet-001/s001-b002-shape-003.png"
    )
    assert make_asset_path(
        sheet_index=1,
        block_id="s001-b001-table",
        asset_kind=AssetKind.RANGE,
        asset_index=1,
    ) == "assets/sheet-001/s001-b001-table-range-001.png"

    with pytest.raises(ValueError):
        make_block_id(sheet_index=0, block_index=1, kind=BlockKind.TABLE)
    with pytest.raises(ValueError):
        make_block_id(sheet_index=1, block_index=1, kind="")
    with pytest.raises(ValueError):
        make_block_id(sheet_index=1, block_index=1, kind="resume")
    with pytest.raises(ValueError):
        make_asset_path(sheet_index=1, block_id=block_id, asset_kind=AssetKind.SHAPE, asset_index=0)


def test_table_header_counts_reject_bool_and_negative_values():
    kwargs = {
        "id": "s001-b001-table",
        "anchor": Rect(sheet="Sheet1", start_row=1, start_col=1, end_row=2, end_col=2, a1="A1:B2"),
        "source": SourceKind.CELLS,
        "rows": [["Name", "Value"], ["A", "1"]],
    }

    with pytest.raises(TypeError, match="header_rows"):
        TableBlock(**kwargs, header_rows=True)
    with pytest.raises(TypeError, match="header_cols"):
        TableBlock(**kwargs, header_cols=False)
    with pytest.raises(ValueError, match="header_rows"):
        TableBlock(**kwargs, header_rows=-1)
    with pytest.raises(ValueError, match="header_cols"):
        TableBlock(**kwargs, header_cols=-1)


def test_models_import_boundary_has_no_external_tool_dependencies():
    code = """
import importlib
import json
import sys
sys.path.insert(0, r'%s')
before = set(sys.modules)
importlib.import_module('excel_semantic_md.models')
after = set(sys.modules) - before
forbidden = [
    'argparse',
    'openpyxl',
    'pythoncom',
    'pywintypes',
    'win32com',
    'win32com.client',
]
print(json.dumps(sorted(name for name in forbidden if name in sys.modules or name in after)))
""" % str(SRC)

    result = subprocess.run(
        [sys.executable, "-c", code],
        check=True,
        capture_output=True,
        text=True,
        env={**os.environ, "PYTHONPATH": str(SRC)},
    )

    assert json.loads(result.stdout) == []


def test_models_module_imports_with_standard_library_only():
    module = importlib.import_module("excel_semantic_md.models")

    assert module.Rect is Rect
