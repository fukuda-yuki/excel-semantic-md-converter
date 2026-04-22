"""Read OOXML visual metadata for Phase 1 inspect output."""

from __future__ import annotations

import posixpath
from dataclasses import dataclass, field
from pathlib import PurePosixPath
from typing import Any
from xml.etree import ElementTree

from excel_semantic_md.excel.workbook_reader import ReadWarning

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"

NS = {
    "a": A_NS,
    "c": C_NS,
    "main": MAIN_NS,
    "pkg": PKG_REL_NS,
    "r": REL_NS,
    "rel": REL_NS,
    "xdr": XDR_NS,
}

CONTENT_TYPE_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CONTENT_TYPE_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CONTENT_TYPE_IMAGE_PNG = "image/png"


@dataclass
class VisualAnchorPoint:
    row: int | None = None
    col: int | None = None
    row_offset_emu: int | None = None
    col_offset_emu: int | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        if self.row is not None:
            data["row"] = self.row
        if self.col is not None:
            data["col"] = self.col
        if self.row_offset_emu is not None:
            data["row_offset_emu"] = self.row_offset_emu
        if self.col_offset_emu is not None:
            data["col_offset_emu"] = self.col_offset_emu
        return data


@dataclass
class VisualAnchor:
    anchor_type: str
    from_point: VisualAnchorPoint | None = None
    to_point: VisualAnchorPoint | None = None
    a1: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"anchor_type": self.anchor_type}
        if self.from_point is not None:
            data["from"] = self.from_point.to_dict()
        if self.to_point is not None:
            data["to"] = self.to_point.to_dict()
        if self.a1 is not None:
            data["a1"] = self.a1
        return data


@dataclass
class VisualSource:
    drawing_part: str
    relationship_id: str | None = None
    target_part: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"drawing_part": self.drawing_part}
        if self.relationship_id is not None:
            data["relationship_id"] = self.relationship_id
        if self.target_part is not None:
            data["target_part"] = self.target_part
        return data


@dataclass
class AssetCandidate:
    kind: str
    source_part: str | None = None
    extension: str | None = None
    content_type: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {"kind": self.kind}
        if self.source_part is not None:
            data["source_part"] = self.source_part
        if self.extension is not None:
            data["extension"] = self.extension
        if self.content_type is not None:
            data["content_type"] = self.content_type
        return data


@dataclass
class ChartSeries:
    name: str | None = None
    categories: list[str] = field(default_factory=list)
    values: list[str] = field(default_factory=list)
    category_ref: str | None = None
    value_ref: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "categories": list(self.categories),
            "values": list(self.values),
        }
        if self.name is not None:
            data["name"] = self.name
        if self.category_ref is not None:
            data["category_ref"] = self.category_ref
        if self.value_ref is not None:
            data["value_ref"] = self.value_ref
        return data


@dataclass
class VisualElement:
    id: str
    kind: str
    anchor: VisualAnchor
    source: VisualSource
    asset_candidate: AssetCandidate
    warnings: list[ReadWarning] = field(default_factory=list)
    text: str | None = None
    shape_type: str | None = None
    alt_text: str | None = None
    title: str | None = None
    series: list[ChartSeries] = field(default_factory=list)
    unknown_kind: str | None = None
    description: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "id": self.id,
            "kind": self.kind,
            "anchor": self.anchor.to_dict(),
            "source": self.source.to_dict(),
            "asset_candidate": self.asset_candidate.to_dict(),
            "warnings": [warning.to_dict() for warning in self.warnings],
        }
        if self.text is not None:
            data["text"] = self.text
        if self.shape_type is not None:
            data["shape_type"] = self.shape_type
        if self.alt_text is not None:
            data["alt_text"] = self.alt_text
        if self.title is not None:
            data["title"] = self.title
        if self.kind == "chart":
            data["series"] = [item.to_dict() for item in self.series]
        if self.unknown_kind is not None:
            data["unknown_kind"] = self.unknown_kind
        if self.description is not None:
            data["description"] = self.description
        return data


@dataclass
class SheetVisualResult:
    sheet_index: int
    name: str
    visuals: list[VisualElement] = field(default_factory=list)
    warnings: list[ReadWarning] = field(default_factory=list)


def read_visual_metadata(path: str | PurePosixPath) -> list[SheetVisualResult]:
    import zipfile

    with zipfile.ZipFile(path) as archive:
        workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        workbook_rels_root = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        content_types = _content_types(archive)
        workbook_relationships = _relationship_targets(workbook_rels_root, "xl/workbook.xml")

        sheets_root = workbook_root.find("main:sheets", NS)
        if sheets_root is None:
            raise KeyError("Workbook XML has no sheets collection.")

        results: list[SheetVisualResult] = []
        for sheet_index, sheet_node in enumerate(sheets_root.findall("main:sheet", NS), start=1):
            state = sheet_node.attrib.get("state", "visible")
            if state != "visible":
                continue
            rel_id = sheet_node.attrib[f"{{{REL_NS}}}id"]
            sheet_path = workbook_relationships[rel_id]
            results.append(
                _read_sheet_visuals(
                    archive=archive,
                    content_types=content_types,
                    sheet_index=sheet_index,
                    sheet_name=sheet_node.attrib["name"],
                    sheet_path=sheet_path,
                )
            )
        return results


def _read_sheet_visuals(
    *,
    archive: Any,
    content_types: dict[str, str],
    sheet_index: int,
    sheet_name: str,
    sheet_path: str,
) -> SheetVisualResult:
    root = ElementTree.fromstring(archive.read(sheet_path))
    result = SheetVisualResult(sheet_index=sheet_index, name=sheet_name)
    drawing_nodes = root.findall("main:drawing", NS)
    if not drawing_nodes:
        return result

    sheet_rels_path = _rels_part_for(sheet_path)
    try:
        sheet_rels_root = ElementTree.fromstring(archive.read(sheet_rels_path))
    except KeyError:
        result.warnings.append(
            ReadWarning(
                code="sheet_drawing_relationships_missing",
                message="Sheet references drawing parts but relationship metadata is missing.",
                details={"sheet": sheet_name, "sheet_part": sheet_path},
            )
        )
        return result

    sheet_relationships = _relationship_targets(sheet_rels_root, sheet_path)
    visual_index = 0
    for drawing_node in drawing_nodes:
        drawing_rel_id = drawing_node.attrib.get(f"{{{REL_NS}}}id")
        if drawing_rel_id is None:
            result.warnings.append(
                ReadWarning(
                    code="sheet_drawing_relationship_id_missing",
                    message="Sheet drawing reference is missing a relationship id.",
                    details={"sheet": sheet_name, "sheet_part": sheet_path},
                )
            )
            continue
        drawing_path = sheet_relationships.get(drawing_rel_id)
        if drawing_path is None:
            result.warnings.append(
                ReadWarning(
                    code="sheet_drawing_target_missing",
                    message="Sheet drawing relationship target is missing.",
                    details={"sheet": sheet_name, "relationship_id": drawing_rel_id},
                )
            )
            continue

        try:
            visuals, drawing_warnings = _read_drawing_part(
                archive=archive,
                content_types=content_types,
                drawing_path=drawing_path,
                sheet_index=sheet_index,
                start_visual_index=visual_index,
            )
        except KeyError:
            result.warnings.append(
                ReadWarning(
                    code="drawing_part_missing",
                    message="Drawing part referenced by the sheet was not found.",
                    details={"sheet": sheet_name, "drawing_part": drawing_path},
                )
            )
            continue
        except ElementTree.ParseError as exc:
            result.warnings.append(
                ReadWarning(
                    code="drawing_part_parse_failed",
                    message="Drawing part could not be parsed.",
                    details={"sheet": sheet_name, "drawing_part": drawing_path, "error": str(exc)},
                )
            )
            continue

        result.visuals.extend(visuals)
        result.warnings.extend(drawing_warnings)
        visual_index += len(visuals)

    return result


def _read_drawing_part(
    *,
    archive: Any,
    content_types: dict[str, str],
    drawing_path: str,
    sheet_index: int,
    start_visual_index: int,
) -> tuple[list[VisualElement], list[ReadWarning]]:
    drawing_root = ElementTree.fromstring(archive.read(drawing_path))
    drawing_relationships = _safe_relationship_targets(archive, _rels_part_for(drawing_path), drawing_path)
    drawing_content_type = content_types.get(drawing_path, CONTENT_TYPE_DRAWING)
    visuals: list[VisualElement] = []
    warnings: list[ReadWarning] = list(drawing_relationships[1])
    current_index = start_visual_index
    for anchor_node in drawing_root:
        anchor_type = _local_name(anchor_node.tag)
        if anchor_type not in {"oneCellAnchor", "twoCellAnchor", "absoluteAnchor"}:
            continue
        current_index += 1
        anchor = _parse_anchor(anchor_node, anchor_type)
        visual = _parse_anchor_visual(
            archive=archive,
            anchor_node=anchor_node,
            anchor=anchor,
            content_types=content_types,
            drawing_content_type=drawing_content_type,
            drawing_path=drawing_path,
            drawing_relationships=drawing_relationships[0],
            sheet_index=sheet_index,
            visual_index=current_index,
        )
        visuals.append(visual)
    return visuals, warnings


def _parse_anchor_visual(
    *,
    archive: Any,
    anchor_node: ElementTree.Element,
    anchor: VisualAnchor,
    content_types: dict[str, str],
    drawing_content_type: str,
    drawing_path: str,
    drawing_relationships: dict[str, str],
    sheet_index: int,
    visual_index: int,
) -> VisualElement:
    shape_node = anchor_node.find("xdr:sp", NS)
    if shape_node is not None:
        return _parse_shape_visual(
            anchor=anchor,
            drawing_content_type=drawing_content_type,
            drawing_path=drawing_path,
            shape_node=shape_node,
            sheet_index=sheet_index,
            visual_index=visual_index,
        )

    pic_node = anchor_node.find("xdr:pic", NS)
    if pic_node is not None:
        return _parse_image_visual(
            anchor=anchor,
            archive=archive,
            content_types=content_types,
            drawing_path=drawing_path,
            drawing_relationships=drawing_relationships,
            pic_node=pic_node,
            sheet_index=sheet_index,
            visual_index=visual_index,
        )

    chart_frame = anchor_node.find("xdr:graphicFrame", NS)
    if chart_frame is not None:
        return _parse_graphic_frame_visual(
            anchor=anchor,
            archive=archive,
            content_types=content_types,
            drawing_content_type=drawing_content_type,
            drawing_path=drawing_path,
            drawing_relationships=drawing_relationships,
            frame_node=chart_frame,
            sheet_index=sheet_index,
            visual_index=visual_index,
        )

    group_shape = anchor_node.find("xdr:grpSp", NS)
    if group_shape is not None:
        return _unknown_visual(
            anchor=anchor,
            drawing_content_type=drawing_content_type,
            drawing_path=drawing_path,
            kind="group_shape",
            description="Group shape is not fully interpreted in Phase 1.",
            sheet_index=sheet_index,
            visual_index=visual_index,
        )

    return _unknown_visual(
        anchor=anchor,
        drawing_content_type=drawing_content_type,
        drawing_path=drawing_path,
        kind="unsupported_anchor_payload",
        description="Drawing anchor payload is not supported in Phase 1.",
        sheet_index=sheet_index,
        visual_index=visual_index,
    )


def _parse_shape_visual(
    *,
    anchor: VisualAnchor,
    drawing_content_type: str,
    drawing_path: str,
    shape_node: ElementTree.Element,
    sheet_index: int,
    visual_index: int,
) -> VisualElement:
    prst_geom = shape_node.find("xdr:spPr/a:prstGeom", NS)
    shape_type = prst_geom.attrib.get("prst") if prst_geom is not None else None
    text = _joined_text(shape_node.findall(".//a:t", NS))
    return VisualElement(
        id=_visual_id(sheet_index, visual_index, "shape"),
        kind="shape",
        anchor=anchor,
        source=VisualSource(drawing_part=drawing_path),
        asset_candidate=AssetCandidate(
            kind="shape",
            source_part=drawing_path,
            extension=_extension(drawing_path),
            content_type=drawing_content_type,
        ),
        shape_type=shape_type,
        text=text if text else None,
    )


def _parse_image_visual(
    *,
    anchor: VisualAnchor,
    archive: Any,
    content_types: dict[str, str],
    drawing_path: str,
    drawing_relationships: dict[str, str],
    pic_node: ElementTree.Element,
    sheet_index: int,
    visual_index: int,
) -> VisualElement:
    nv_pic_pr = pic_node.find("xdr:nvPicPr/xdr:cNvPr", NS)
    alt_text = None if nv_pic_pr is None else nv_pic_pr.attrib.get("descr")
    blip = pic_node.find("xdr:blipFill/a:blip", NS)
    rel_id = None if blip is None else blip.attrib.get(f"{{{REL_NS}}}embed")
    target_part = None if rel_id is None else drawing_relationships.get(rel_id)
    warnings: list[ReadWarning] = []
    if rel_id is not None and target_part is None:
        warnings.append(
            ReadWarning(
                code="image_target_missing",
                message="Embedded image relationship target is missing.",
                details={"drawing_part": drawing_path, "relationship_id": rel_id},
            )
        )
    if target_part is not None and not _part_exists(archive, target_part):
        warnings.append(
            ReadWarning(
                code="image_part_missing",
                message="Embedded image part was not found in the workbook package.",
                details={"drawing_part": drawing_path, "target_part": target_part},
            )
        )
    return VisualElement(
        id=_visual_id(sheet_index, visual_index, "image"),
        kind="image",
        anchor=anchor,
        source=VisualSource(drawing_part=drawing_path, relationship_id=rel_id, target_part=target_part),
        asset_candidate=AssetCandidate(
            kind="image",
            source_part=target_part,
            extension=_extension(target_part),
            content_type=content_types.get(target_part) if target_part is not None else None,
        ),
        warnings=warnings,
        alt_text=alt_text,
    )


def _parse_graphic_frame_visual(
    *,
    anchor: VisualAnchor,
    archive: Any,
    content_types: dict[str, str],
    drawing_content_type: str,
    drawing_path: str,
    drawing_relationships: dict[str, str],
    frame_node: ElementTree.Element,
    sheet_index: int,
    visual_index: int,
) -> VisualElement:
    graphic_data = frame_node.find("a:graphic/a:graphicData", NS)
    if graphic_data is None:
        return _unknown_visual(
            anchor=anchor,
            drawing_content_type=drawing_content_type,
            drawing_path=drawing_path,
            kind="graphic_frame",
            description="Graphic frame does not contain graphicData.",
            sheet_index=sheet_index,
            visual_index=visual_index,
        )

    chart_node = graphic_data.find("c:chart", NS)
    if chart_node is None:
        uri = graphic_data.attrib.get("uri", "")
        unknown_kind = "smartart" if "diagram" in uri else "graphic_frame"
        return _unknown_visual(
            anchor=anchor,
            drawing_content_type=drawing_content_type,
            drawing_path=drawing_path,
            kind=unknown_kind,
            description=f"Graphic frame URI is not supported in Phase 1: {uri or 'unknown'}",
            sheet_index=sheet_index,
            visual_index=visual_index,
        )

    rel_id = chart_node.attrib.get(f"{{{REL_NS}}}id")
    target_part = None if rel_id is None else drawing_relationships.get(rel_id)
    warnings: list[ReadWarning] = []
    title: str | None = None
    series: list[ChartSeries] = []
    if rel_id is None:
        warnings.append(
            ReadWarning(
                code="chart_relationship_id_missing",
                message="Chart graphic frame is missing a relationship id.",
                details={"drawing_part": drawing_path},
            )
        )
    elif target_part is None:
        warnings.append(
            ReadWarning(
                code="chart_target_missing",
                message="Chart relationship target is missing.",
                details={"drawing_part": drawing_path, "relationship_id": rel_id},
            )
        )
    elif not _part_exists(archive, target_part):
        warnings.append(
            ReadWarning(
                code="chart_part_missing",
                message="Chart part was not found in the workbook package.",
                details={"drawing_part": drawing_path, "target_part": target_part},
            )
        )
    else:
        try:
            chart_root = ElementTree.fromstring(archive.read(target_part))
        except ElementTree.ParseError as exc:
            warnings.append(
                ReadWarning(
                    code="chart_part_parse_failed",
                    message="Chart part could not be parsed.",
                    details={"target_part": target_part, "error": str(exc)},
                )
            )
        else:
            title = _chart_title(chart_root)
            series = _chart_series(chart_root)

    return VisualElement(
        id=_visual_id(sheet_index, visual_index, "chart"),
        kind="chart",
        anchor=anchor,
        source=VisualSource(drawing_part=drawing_path, relationship_id=rel_id, target_part=target_part),
        asset_candidate=AssetCandidate(
            kind="chart",
            source_part=target_part,
            extension=_extension(target_part),
            content_type=content_types.get(target_part) if target_part is not None else None,
        ),
        warnings=warnings,
        title=title,
        series=series,
    )


def _unknown_visual(
    *,
    anchor: VisualAnchor,
    drawing_content_type: str,
    drawing_path: str,
    kind: str,
    description: str,
    sheet_index: int,
    visual_index: int,
) -> VisualElement:
    return VisualElement(
        id=_visual_id(sheet_index, visual_index, "unknown"),
        kind="unknown",
        anchor=anchor,
        source=VisualSource(drawing_part=drawing_path),
        asset_candidate=AssetCandidate(
            kind="unknown",
            source_part=drawing_path,
            extension=_extension(drawing_path),
            content_type=drawing_content_type,
        ),
        warnings=[
            ReadWarning(
                code="unsupported_visual_element",
                message=description,
                details={"drawing_part": drawing_path, "unknown_kind": kind},
            )
        ],
        unknown_kind=kind,
        description=description,
    )


def _parse_anchor(anchor_node: ElementTree.Element, anchor_type: str) -> VisualAnchor:
    if anchor_type == "absoluteAnchor":
        return VisualAnchor(anchor_type=anchor_type)

    from_node = anchor_node.find("xdr:from", NS)
    to_node = anchor_node.find("xdr:to", NS)
    from_point = _anchor_point(from_node)
    to_point = _anchor_point(to_node)
    return VisualAnchor(
        anchor_type=anchor_type,
        from_point=from_point,
        to_point=to_point,
        a1=_anchor_a1(from_point, to_point),
    )


def _anchor_point(node: ElementTree.Element | None) -> VisualAnchorPoint | None:
    if node is None:
        return None
    row_text = node.findtext("xdr:row", default=None, namespaces=NS)
    col_text = node.findtext("xdr:col", default=None, namespaces=NS)
    row_offset_text = node.findtext("xdr:rowOff", default=None, namespaces=NS)
    col_offset_text = node.findtext("xdr:colOff", default=None, namespaces=NS)
    return VisualAnchorPoint(
        row=None if row_text is None else int(row_text) + 1,
        col=None if col_text is None else int(col_text) + 1,
        row_offset_emu=None if row_offset_text is None else int(row_offset_text),
        col_offset_emu=None if col_offset_text is None else int(col_offset_text),
    )


def _anchor_a1(
    from_point: VisualAnchorPoint | None,
    to_point: VisualAnchorPoint | None,
) -> str | None:
    if from_point is None or from_point.row is None or from_point.col is None:
        return None
    if to_point is None or to_point.row is None or to_point.col is None:
        return _cell_ref(from_point.row, from_point.col)
    return f"{_cell_ref(from_point.row, from_point.col)}:{_cell_ref(to_point.row, to_point.col)}"


def _cell_ref(row: int, col: int) -> str:
    letters: list[str] = []
    remaining = col
    while remaining > 0:
        remaining, offset = divmod(remaining - 1, 26)
        letters.append(chr(ord("A") + offset))
    return "".join(reversed(letters)) + str(row)


def _chart_title(root: ElementTree.Element) -> str | None:
    title_node = root.find(".//c:chart/c:title", NS)
    if title_node is None:
        return None
    text = _joined_text(title_node.findall(".//a:t", NS))
    if text:
        return text
    values = _joined_text(title_node.findall(".//c:v", NS))
    return values or None


def _chart_series(root: ElementTree.Element) -> list[ChartSeries]:
    items: list[ChartSeries] = []
    for series_node in root.findall(".//c:ser", NS):
        items.append(
            ChartSeries(
                name=_series_name(series_node),
                categories=_series_values(series_node.find("c:cat", NS)),
                values=_series_values(series_node.find("c:val", NS)),
                category_ref=_series_ref(series_node.find("c:cat", NS)),
                value_ref=_series_ref(series_node.find("c:val", NS)),
            )
        )
    return items


def _series_name(series_node: ElementTree.Element) -> str | None:
    text_node = series_node.find("c:tx", NS)
    if text_node is None:
        return None
    rich_text = _joined_text(text_node.findall(".//a:t", NS))
    if rich_text:
        return rich_text
    value = text_node.findtext(".//c:v", default=None, namespaces=NS)
    if value is not None:
        return value
    formula = text_node.findtext(".//c:f", default=None, namespaces=NS)
    return formula


def _series_values(node: ElementTree.Element | None) -> list[str]:
    if node is None:
        return []
    cache = node.find(".//c:strCache", NS)
    if cache is None:
        cache = node.find(".//c:numCache", NS)
    if cache is None:
        literal = node.find(".//c:strLit", NS)
        if literal is None:
            literal = node.find(".//c:numLit", NS)
        cache = literal
    if cache is None:
        return []
    values: list[str] = []
    for point in cache.findall("c:pt", NS):
        value = point.findtext("c:v", default="", namespaces=NS)
        values.append(value)
    return values


def _series_ref(node: ElementTree.Element | None) -> str | None:
    if node is None:
        return None
    formula = node.findtext(".//c:f", default=None, namespaces=NS)
    return formula


def _joined_text(nodes: list[ElementTree.Element]) -> str:
    return "\n".join(text for text in (node.text for node in nodes) if text)


def _content_types(archive: Any) -> dict[str, str]:
    root = ElementTree.fromstring(archive.read("[Content_Types].xml"))
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for node in root:
        local = _local_name(node.tag)
        if local == "Default":
            defaults[node.attrib["Extension"].lower()] = node.attrib["ContentType"]
        elif local == "Override":
            overrides[node.attrib["PartName"].lstrip("/")] = node.attrib["ContentType"]
    result = dict(overrides)
    for part_name in archive.namelist():
        if part_name in result:
            continue
        extension = PurePosixPath(part_name).suffix.lstrip(".").lower()
        if extension in defaults:
            result[part_name] = defaults[extension]
    return result


def _relationship_targets(rels_root: ElementTree.Element, source_part: str) -> dict[str, str]:
    targets: dict[str, str] = {}
    for rel_node in rels_root.findall("pkg:Relationship", NS):
        rel_id = rel_node.attrib["Id"]
        target = rel_node.attrib["Target"]
        targets[rel_id] = _resolve_target(source_part, target)
    return targets


def _safe_relationship_targets(
    archive: Any,
    rels_path: str,
    source_part: str,
) -> tuple[dict[str, str], list[ReadWarning]]:
    try:
        rels_root = ElementTree.fromstring(archive.read(rels_path))
    except KeyError:
        return {}, []
    except ElementTree.ParseError as exc:
        return (
            {},
            [
                ReadWarning(
                    code="drawing_relationships_parse_failed",
                    message="Drawing relationships part could not be parsed.",
                    details={"relationships_part": rels_path, "error": str(exc)},
                )
            ],
        )
    return _relationship_targets(rels_root, source_part), []


def _resolve_target(source_part: str, target: str) -> str:
    normalized_target = target.replace("\\", "/")
    if normalized_target.startswith("/"):
        return normalized_target.lstrip("/")
    base_dir = str(PurePosixPath(source_part).parent)
    joined = normalized_target if base_dir == "." else f"{base_dir}/{normalized_target}"
    return posixpath.normpath(joined)


def _rels_part_for(part_path: str) -> str:
    pure = PurePosixPath(part_path)
    return str(pure.parent / "_rels" / f"{pure.name}.rels")


def _extension(part_path: str | None) -> str | None:
    if part_path is None:
        return None
    suffix = PurePosixPath(part_path).suffix
    return suffix or None


def _part_exists(archive: Any, part_path: str) -> bool:
    try:
        archive.getinfo(part_path)
        return True
    except KeyError:
        return False


def _visual_id(sheet_index: int, visual_index: int, kind: str) -> str:
    return f"s{sheet_index:03d}-v{visual_index:03d}-{kind}"


def _local_name(tag: str) -> str:
    if "}" not in tag:
        return tag
    return tag.split("}", 1)[1]


__all__ = ["read_visual_metadata", "SheetVisualResult", "VisualElement"]
