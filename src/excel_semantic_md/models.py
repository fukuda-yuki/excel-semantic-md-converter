"""Core Phase 1 data models for workbook extraction and manifests."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, ClassVar


SCHEMA_VERSION = "phase1.0"


class BlockKind(Enum):
    HEADING = "heading"
    PARAGRAPH = "paragraph"
    TABLE = "table"
    SHAPE = "shape"
    IMAGE = "image"
    CHART = "chart"
    NOTE = "note"
    UNKNOWN = "unknown"


class SourceKind(Enum):
    CELLS = "cells"
    SHAPE = "shape"
    IMAGE = "image"
    CHART = "chart"
    UNKNOWN = "unknown"


class AssetKind(Enum):
    RANGE = "range"
    SHAPE = "shape"
    IMAGE = "image"
    CHART = "chart"
    UNKNOWN = "unknown"


class AssetRole(Enum):
    MARKDOWN = "markdown"
    LLM_ATTACHMENT = "llm_attachment"
    RENDER_ARTIFACT = "render_artifact"


def _enum_value(value: Enum | str) -> str:
    if isinstance(value, Enum):
        return str(value.value)
    return str(value)


def _enum_from(enum_type: type[Enum], value: Enum | str) -> Enum:
    if isinstance(value, enum_type):
        return value
    return enum_type(str(value))


def make_block_id(sheet_index: int, block_index: int, kind: BlockKind | str) -> str:
    """Return the stable Phase 1 block ID."""

    _validate_positive_int("sheet_index", sheet_index)
    _validate_positive_int("block_index", block_index)
    kind_value = _enum_from(BlockKind, kind).value
    return f"s{sheet_index:03d}-b{block_index:03d}-{kind_value}"


def make_asset_path(
    sheet_index: int,
    block_id: str,
    asset_kind: AssetKind | str,
    asset_index: int,
) -> str:
    """Return the stable Phase 1 asset path."""

    _validate_positive_int("sheet_index", sheet_index)
    _validate_positive_int("asset_index", asset_index)
    if not block_id:
        raise ValueError("block_id must not be empty")
    asset_kind_value = _enum_from(AssetKind, asset_kind).value
    filename_stem = block_id
    if not filename_stem.endswith(f"-{asset_kind_value}"):
        filename_stem = f"{filename_stem}-{asset_kind_value}"
    return f"assets/sheet-{sheet_index:03d}/{filename_stem}-{asset_index:03d}.png"


def _validate_positive_int(name: str, value: int) -> None:
    if not isinstance(value, int) or isinstance(value, bool):
        raise TypeError(f"{name} must be an integer")
    if value < 1:
        raise ValueError(f"{name} must be 1 or greater")


def _validate_non_negative_int(name: str, value: int) -> None:
    if not isinstance(value, int) or isinstance(value, bool):
        raise TypeError(f"{name} must be an integer")
    if value < 0:
        raise ValueError(f"{name} must be 0 or greater")


@dataclass
class Rect:
    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    a1: str

    def __post_init__(self) -> None:
        if not self.sheet:
            raise ValueError("sheet must not be empty")
        if not self.a1:
            raise ValueError("a1 must not be empty")
        _validate_positive_int("start_row", self.start_row)
        _validate_positive_int("start_col", self.start_col)
        _validate_positive_int("end_row", self.end_row)
        _validate_positive_int("end_col", self.end_col)
        if self.start_row > self.end_row:
            raise ValueError("start_row must be less than or equal to end_row")
        if self.start_col > self.end_col:
            raise ValueError("start_col must be less than or equal to end_col")

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "sheet": self.sheet,
            "start_row": self.start_row,
            "start_col": self.start_col,
            "end_row": self.end_row,
            "end_col": self.end_col,
            "a1": self.a1,
        }
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Rect":
        return cls(
            sheet=data["sheet"],
            start_row=data["start_row"],
            start_col=data["start_col"],
            end_row=data["end_row"],
            end_col=data["end_col"],
            a1=data["a1"],
        )


@dataclass
class WarningInfo:
    code: str
    message: str
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {"code": self.code, "message": self.message, "details": dict(self.details)}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "WarningInfo":
        return cls(code=data["code"], message=data["message"], details=dict(data.get("details", {})))


@dataclass
class UnknownInfo:
    kind: str
    description: str
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {"kind": self.kind, "description": self.description, "details": dict(self.details)}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "UnknownInfo":
        return cls(
            kind=data["kind"],
            description=data["description"],
            details=dict(data.get("details", {})),
        )


@dataclass
class FailureInfo:
    stage: str
    message: str
    details: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {"stage": self.stage, "message": self.message, "details": dict(self.details)}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "FailureInfo":
        return cls(stage=data["stage"], message=data["message"], details=dict(data.get("details", {})))


@dataclass
class AssetRef:
    path: str
    kind: AssetKind
    role: AssetRole
    description: str | None = None

    def __post_init__(self) -> None:
        if not self.path:
            raise ValueError("path must not be empty")
        self.kind = _enum_from(AssetKind, self.kind)  # type: ignore[assignment]
        self.role = _enum_from(AssetRole, self.role)  # type: ignore[assignment]

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "path": self.path,
            "kind": self.kind.value,
            "role": self.role.value,
        }
        if self.description is not None:
            data["description"] = self.description
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AssetRef":
        return cls(
            path=data["path"],
            kind=AssetKind(data["kind"]),
            role=AssetRole(data["role"]),
            description=data.get("description"),
        )


@dataclass
class Block:
    id: str
    anchor: Rect
    source: SourceKind
    visual_id: str | None = None
    related_block_id: str | None = None
    assets: list[AssetRef] = field(default_factory=list)
    warnings: list[WarningInfo] = field(default_factory=list)

    kind: ClassVar[BlockKind] = BlockKind.UNKNOWN

    def __post_init__(self) -> None:
        if not self.id:
            raise ValueError("id must not be empty")
        if not isinstance(self.anchor, Rect):
            self.anchor = Rect.from_dict(self.anchor)  # type: ignore[assignment]
        self.source = _enum_from(SourceKind, self.source)  # type: ignore[assignment]
        self.assets = [asset if isinstance(asset, AssetRef) else AssetRef.from_dict(asset) for asset in self.assets]
        self.warnings = [
            warning if isinstance(warning, WarningInfo) else WarningInfo.from_dict(warning)
            for warning in self.warnings
        ]

    def to_dict(self) -> dict[str, Any]:
        data = self._base_dict()
        data.update(self._specific_dict())
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Block":
        kind = BlockKind(data["kind"])
        block_type = _BLOCK_TYPES[kind]
        return block_type._from_dict(data)

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "Block":
        return cls(**cls._base_kwargs(data))

    @classmethod
    def _base_kwargs(cls, data: dict[str, Any]) -> dict[str, Any]:
        return {
            "id": data["id"],
            "anchor": Rect.from_dict(data["anchor"]),
            "source": SourceKind(data["source"]),
            "visual_id": data.get("visual_id"),
            "related_block_id": data.get("related_block_id"),
            "assets": [AssetRef.from_dict(item) for item in data.get("assets", [])],
            "warnings": [WarningInfo.from_dict(item) for item in data.get("warnings", [])],
        }

    def _base_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "kind": self.kind.value,
            "anchor": self.anchor.to_dict(),
            "source": self.source.value,
            "visual_id": self.visual_id,
            "related_block_id": self.related_block_id,
            "assets": [asset.to_dict() for asset in self.assets],
            "warnings": [warning.to_dict() for warning in self.warnings],
        }

    def _specific_dict(self) -> dict[str, Any]:
        return {}


@dataclass
class HeadingBlock(Block):
    text: str = ""
    level: int = 1

    kind: ClassVar[BlockKind] = BlockKind.HEADING

    def __post_init__(self) -> None:
        super().__post_init__()
        _validate_positive_int("level", self.level)

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "HeadingBlock":
        return cls(**cls._base_kwargs(data), text=data.get("text", ""), level=data.get("level", 1))

    def _specific_dict(self) -> dict[str, Any]:
        return {"text": self.text, "level": self.level}


@dataclass
class ParagraphBlock(Block):
    text: str = ""

    kind: ClassVar[BlockKind] = BlockKind.PARAGRAPH

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "ParagraphBlock":
        return cls(**cls._base_kwargs(data), text=data.get("text", ""))

    def _specific_dict(self) -> dict[str, Any]:
        return {"text": self.text}


@dataclass
class TableBlock(Block):
    rows: list[list[str]] = field(default_factory=list)
    header_rows: int = 0
    header_cols: int = 0

    kind: ClassVar[BlockKind] = BlockKind.TABLE

    def __post_init__(self) -> None:
        super().__post_init__()
        _validate_non_negative_int("header_rows", self.header_rows)
        _validate_non_negative_int("header_cols", self.header_cols)

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "TableBlock":
        return cls(
            **cls._base_kwargs(data),
            rows=[list(row) for row in data.get("rows", [])],
            header_rows=data.get("header_rows", 0),
            header_cols=data.get("header_cols", 0),
        )

    def _specific_dict(self) -> dict[str, Any]:
        return {
            "rows": [list(row) for row in self.rows],
            "header_rows": self.header_rows,
            "header_cols": self.header_cols,
        }


@dataclass
class ShapeBlock(Block):
    text: str = ""
    shape_type: str | None = None

    kind: ClassVar[BlockKind] = BlockKind.SHAPE

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "ShapeBlock":
        return cls(
            **cls._base_kwargs(data),
            text=data.get("text", ""),
            shape_type=data.get("shape_type"),
        )

    def _specific_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"text": self.text}
        if self.shape_type is not None:
            data["shape_type"] = self.shape_type
        return data


@dataclass
class ImageBlock(Block):
    alt_text: str | None = None

    kind: ClassVar[BlockKind] = BlockKind.IMAGE

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "ImageBlock":
        return cls(**cls._base_kwargs(data), alt_text=data.get("alt_text"))

    def _specific_dict(self) -> dict[str, Any]:
        return {} if self.alt_text is None else {"alt_text": self.alt_text}


@dataclass
class ChartBlock(Block):
    title: str | None = None
    metadata: dict[str, Any] = field(default_factory=dict)

    kind: ClassVar[BlockKind] = BlockKind.CHART

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "ChartBlock":
        return cls(
            **cls._base_kwargs(data),
            title=data.get("title"),
            metadata=dict(data.get("metadata", {})),
        )

    def _specific_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"metadata": dict(self.metadata)}
        if self.title is not None:
            data["title"] = self.title
        return data


@dataclass
class NoteBlock(Block):
    text: str = ""

    kind: ClassVar[BlockKind] = BlockKind.NOTE

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "NoteBlock":
        return cls(**cls._base_kwargs(data), text=data.get("text", ""))

    def _specific_dict(self) -> dict[str, Any]:
        return {"text": self.text}


@dataclass
class UnknownBlock(Block):
    unknown: UnknownInfo | None = None

    kind: ClassVar[BlockKind] = BlockKind.UNKNOWN

    def __post_init__(self) -> None:
        super().__post_init__()
        if self.unknown is not None and not isinstance(self.unknown, UnknownInfo):
            self.unknown = UnknownInfo.from_dict(self.unknown)  # type: ignore[assignment]

    @classmethod
    def _from_dict(cls, data: dict[str, Any]) -> "UnknownBlock":
        unknown = data.get("unknown")
        return cls(
            **cls._base_kwargs(data),
            unknown=UnknownInfo.from_dict(unknown) if unknown is not None else None,
        )

    def _specific_dict(self) -> dict[str, Any]:
        return {} if self.unknown is None else {"unknown": self.unknown.to_dict()}


_BLOCK_TYPES: dict[BlockKind, type[Block]] = {
    BlockKind.HEADING: HeadingBlock,
    BlockKind.PARAGRAPH: ParagraphBlock,
    BlockKind.TABLE: TableBlock,
    BlockKind.SHAPE: ShapeBlock,
    BlockKind.IMAGE: ImageBlock,
    BlockKind.CHART: ChartBlock,
    BlockKind.NOTE: NoteBlock,
    BlockKind.UNKNOWN: UnknownBlock,
}


@dataclass
class SheetModel:
    sheet_index: int
    name: str
    blocks: list[Block] = field(default_factory=list)
    failures: list[FailureInfo] = field(default_factory=list)
    warnings: list[WarningInfo] = field(default_factory=list)

    def __post_init__(self) -> None:
        _validate_positive_int("sheet_index", self.sheet_index)
        if not self.name:
            raise ValueError("name must not be empty")
        self.blocks = [block if isinstance(block, Block) else Block.from_dict(block) for block in self.blocks]
        self.failures = [
            failure if isinstance(failure, FailureInfo) else FailureInfo.from_dict(failure)
            for failure in self.failures
        ]
        self.warnings = [
            warning if isinstance(warning, WarningInfo) else WarningInfo.from_dict(warning)
            for warning in self.warnings
        ]

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet_index": self.sheet_index,
            "name": self.name,
            "blocks": [block.to_dict() for block in self.blocks],
            "failures": [failure.to_dict() for failure in self.failures],
            "warnings": [warning.to_dict() for warning in self.warnings],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "SheetModel":
        return cls(
            sheet_index=data["sheet_index"],
            name=data["name"],
            blocks=[Block.from_dict(item) for item in data.get("blocks", [])],
            failures=[FailureInfo.from_dict(item) for item in data.get("failures", [])],
            warnings=[WarningInfo.from_dict(item) for item in data.get("warnings", [])],
        )


@dataclass
class WorkbookModel:
    sheets: list[SheetModel] = field(default_factory=list)
    input_file_name: str | None = None
    schema_version: str = SCHEMA_VERSION

    def __post_init__(self) -> None:
        self.sheets = [sheet if isinstance(sheet, SheetModel) else SheetModel.from_dict(sheet) for sheet in self.sheets]

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "schema_version": self.schema_version,
            "sheets": [sheet.to_dict() for sheet in self.sheets],
        }
        if self.input_file_name is not None:
            data["input_file_name"] = self.input_file_name
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "WorkbookModel":
        return cls(
            sheets=[SheetModel.from_dict(item) for item in data.get("sheets", [])],
            input_file_name=data.get("input_file_name"),
            schema_version=data.get("schema_version", SCHEMA_VERSION),
        )


__all__ = [
    "SCHEMA_VERSION",
    "AssetKind",
    "AssetRef",
    "AssetRole",
    "Block",
    "BlockKind",
    "ChartBlock",
    "FailureInfo",
    "HeadingBlock",
    "ImageBlock",
    "NoteBlock",
    "ParagraphBlock",
    "Rect",
    "ShapeBlock",
    "SheetModel",
    "SourceKind",
    "TableBlock",
    "UnknownBlock",
    "UnknownInfo",
    "WarningInfo",
    "WorkbookModel",
    "make_asset_path",
    "make_block_id",
]
