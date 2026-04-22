"""Render planning and result models for Excel COM confirmation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from excel_semantic_md.models import AssetRole, Block, FailureInfo, Rect, WarningInfo


@dataclass(frozen=True)
class RenderPlanItem:
    block: Block
    kind: str
    role: AssetRole
    source: str
    target_part: str | None = None


@dataclass
class RenderArtifact:
    block_id: str
    visual_id: str | None
    related_block_id: str | None
    kind: str
    role: str
    path: str
    source: str
    anchor: Rect

    def to_dict(self) -> dict[str, Any]:
        return {
            "block_id": self.block_id,
            "visual_id": self.visual_id,
            "related_block_id": self.related_block_id,
            "kind": self.kind,
            "role": self.role,
            "path": self.path,
            "source": self.source,
            "anchor": self.anchor.to_dict(),
        }


@dataclass
class RenderSheetResult:
    input_file_name: str
    sheet_name: str
    temp_dir: str
    artifacts: list[RenderArtifact] = field(default_factory=list)
    warnings: list[WarningInfo] = field(default_factory=list)
    failures: list[FailureInfo] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "input_file_name": self.input_file_name,
            "sheet_name": self.sheet_name,
            "temp_dir": self.temp_dir,
            "artifacts": [artifact.to_dict() for artifact in self.artifacts],
            "warnings": [warning.to_dict() for warning in self.warnings],
            "failures": [failure.to_dict() for failure in self.failures],
        }

