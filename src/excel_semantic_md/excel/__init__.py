"""Workbook extraction layer."""

from excel_semantic_md.excel.block_detector import detect_blocks
from excel_semantic_md.excel.ooxml_visual_reader import read_visual_metadata
from excel_semantic_md.excel.visual_linker import link_visuals
from excel_semantic_md.excel.workbook_reader import (
    CellReadValue,
    MergedRange,
    ReadFailure,
    ReadWarning,
    SheetReadResult,
    WorkbookReadResult,
    read_workbook,
)

__all__ = [
    "CellReadValue",
    "detect_blocks",
    "MergedRange",
    "read_visual_metadata",
    "link_visuals",
    "ReadFailure",
    "ReadWarning",
    "SheetReadResult",
    "WorkbookReadResult",
    "read_workbook",
]
