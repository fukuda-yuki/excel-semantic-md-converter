"""LLM-specific models for sheet-level Copilot integration."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from excel_semantic_md.models import FailureInfo


def _validate_non_negative_int(name: str, value: int | None) -> None:
    if value is None:
        return
    if not isinstance(value, int) or isinstance(value, bool):
        raise TypeError(f"{name} must be an integer")
    if value < 0:
        raise ValueError(f"{name} must be 0 or greater")


@dataclass
class LlmRunOptions:
    model: str | None = None
    vision_model: str | None = None
    max_images_per_sheet: int | None = None

    def __post_init__(self) -> None:
        _validate_non_negative_int("max_images_per_sheet", self.max_images_per_sheet)


@dataclass(frozen=True)
class LlmAttachment:
    path: str
    block_id: str | None
    related_block_id: str | None
    kind: str
    source: str
    priority: int

    def __post_init__(self) -> None:
        if not self.path:
            raise ValueError("path must not be empty")
        _validate_non_negative_int("priority", self.priority)

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "block_id": self.block_id,
            "related_block_id": self.related_block_id,
            "kind": self.kind,
            "source": self.source,
            "priority": self.priority,
        }


@dataclass
class LlmInput:
    sheet_name: str
    blocks: list[dict[str, Any]] = field(default_factory=list)
    assets: list[dict[str, Any]] = field(default_factory=list)
    instructions: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.sheet_name:
            raise ValueError("sheet_name must not be empty")

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheetName": self.sheet_name,
            "blocks": [dict(block) for block in self.blocks],
            "assets": [dict(asset) for asset in self.assets],
            "instructions": dict(self.instructions),
        }


@dataclass
class LlmResponse:
    sheet_summary: str
    sections: list[Any]
    figures: list[Any]
    unknowns: list[Any]
    markdown: str
    raw: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet_summary": self.sheet_summary,
            "sections": list(self.sections),
            "figures": list(self.figures),
            "unknowns": list(self.unknowns),
            "markdown": self.markdown,
            "raw": dict(self.raw),
        }


@dataclass
class LlmRunResult:
    status: str
    attempts: int
    response: LlmResponse | None = None
    failure: FailureInfo | None = None

    def __post_init__(self) -> None:
        if self.status not in {"succeeded", "failed"}:
            raise ValueError("status must be either 'succeeded' or 'failed'")
        if not isinstance(self.attempts, int) or isinstance(self.attempts, bool):
            raise TypeError("attempts must be an integer")
        if self.attempts < 1:
            raise ValueError("attempts must be 1 or greater")
        if self.failure is not None and not isinstance(self.failure, FailureInfo):
            self.failure = FailureInfo.from_dict(self.failure)  # type: ignore[assignment]

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "status": self.status,
            "attempts": self.attempts,
            "response": None if self.response is None else self.response.to_dict(),
            "failure": None if self.failure is None else self.failure.to_dict(),
        }
        return data
