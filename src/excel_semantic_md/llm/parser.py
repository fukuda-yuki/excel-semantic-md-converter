"""Parser and validator for LLM response JSON."""

from __future__ import annotations

import json
import re
from typing import Any

from excel_semantic_md.llm.models import LlmResponse


_FENCED_JSON_PATTERN = re.compile(r"```(?:json)?\s*(?P<body>.*?)\s*```", re.IGNORECASE | re.DOTALL)


def parse_llm_response(raw_text: str) -> LlmResponse:
    if not raw_text or not raw_text.strip():
        raise ValueError("LLM response text must not be empty")

    payload = _parse_json_payload(_extract_json_text(raw_text))
    _validate_payload(payload)
    return LlmResponse(
        sheet_summary=payload["sheet_summary"],
        sections=list(payload["sections"]),
        figures=list(payload["figures"]),
        unknowns=list(payload["unknowns"]),
        markdown=payload["markdown"],
        raw=dict(payload),
    )


def _extract_json_text(raw_text: str) -> str:
    stripped = raw_text.strip()
    fenced = _FENCED_JSON_PATTERN.fullmatch(stripped)
    if fenced is not None:
        return fenced.group("body").strip()

    fenced = _FENCED_JSON_PATTERN.search(stripped)
    if fenced is not None:
        return fenced.group("body").strip()
    return stripped


def _parse_json_payload(raw_json: str) -> dict[str, Any]:
    try:
        payload = json.loads(raw_json)
    except json.JSONDecodeError as exc:
        raise ValueError(f"LLM response is not valid JSON: {exc}") from exc
    if not isinstance(payload, dict):
        raise ValueError("LLM response JSON must be an object")
    return payload


def _validate_payload(payload: dict[str, Any]) -> None:
    required = {"sheet_summary", "sections", "figures", "unknowns", "markdown"}
    missing = sorted(required - payload.keys())
    if missing:
        raise ValueError(f"LLM response is missing required keys: {', '.join(missing)}")

    if not isinstance(payload["sheet_summary"], str):
        raise ValueError("LLM response field 'sheet_summary' must be a string")
    if not isinstance(payload["sections"], list):
        raise ValueError("LLM response field 'sections' must be a list")
    if not isinstance(payload["figures"], list):
        raise ValueError("LLM response field 'figures' must be a list")
    if not isinstance(payload["unknowns"], list):
        raise ValueError("LLM response field 'unknowns' must be a list")
    if not isinstance(payload["markdown"], str):
        raise ValueError("LLM response field 'markdown' must be a string")
    if not payload["markdown"].strip():
        raise ValueError("LLM response field 'markdown' must not be empty")
