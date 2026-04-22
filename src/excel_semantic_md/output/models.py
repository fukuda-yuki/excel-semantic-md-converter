"""Internal models for convert output generation."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from excel_semantic_md.llm.models import LlmRunResult
from excel_semantic_md.models import Block, FailureInfo, Rect, SheetModel, WarningInfo
from excel_semantic_md.render.types import RenderSheetResult


@dataclass
class PublishedAsset:
    sheet_index: int
    sheet_name: str
    block_id: str
    visual_id: str | None
    related_block_id: str | None
    kind: str
    role: str
    source: str
    path: str
    anchor: Rect

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet_index": self.sheet_index,
            "sheet_name": self.sheet_name,
            "block_id": self.block_id,
            "visual_id": self.visual_id,
            "related_block_id": self.related_block_id,
            "kind": self.kind,
            "role": self.role,
            "source": self.source,
            "path": self.path,
            "anchor": self.anchor.to_dict(),
        }


@dataclass
class ConvertSheetResult:
    sheet: SheetModel
    status: str
    warnings: list[WarningInfo] = field(default_factory=list)
    failures: list[FailureInfo] = field(default_factory=list)
    markdown: str | None = None
    render_plan_payload: dict[str, Any] | None = None
    render_result: RenderSheetResult | None = None
    llm_input_payload: dict[str, Any] | None = None
    llm_result: LlmRunResult | None = None
    assets: list[PublishedAsset] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.status not in {"succeeded", "failed"}:
            raise ValueError("status must be either 'succeeded' or 'failed'")


@dataclass
class ConvertResult:
    input_file_name: str
    schema_version: str
    generated_at: str
    command_options: dict[str, Any]
    output_dir: Path
    workbook_extraction_payload: dict[str, Any]
    block_detection_payload: dict[str, Any]
    linked_workbook_payload: dict[str, Any]
    sheets: list[ConvertSheetResult] = field(default_factory=list)

    @property
    def failed_sheet_count(self) -> int:
        return sum(1 for sheet in self.sheets if sheet.status == "failed")

    @property
    def has_failures(self) -> bool:
        return self.failed_sheet_count > 0

    @property
    def blocks(self) -> list[Block]:
        return [block for sheet in self.sheets for block in sheet.sheet.blocks]


@dataclass(frozen=True)
class ConvertOutputFiles:
    result_markdown: Path
    manifest_json: Path
    assets_dir: Path
    debug_dir: Path | None = None
