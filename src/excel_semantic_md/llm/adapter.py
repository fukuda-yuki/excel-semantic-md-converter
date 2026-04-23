"""GitHub Copilot SDK adapter for sheet-level LLM execution."""

from __future__ import annotations

import asyncio
from typing import Any

from excel_semantic_md.llm.builders import build_llm_attachments, build_llm_input
from excel_semantic_md.llm.models import LlmRunOptions, LlmRunResult
from excel_semantic_md.llm.parser import parse_llm_response
from excel_semantic_md.llm.prompt import build_sheet_prompt
from excel_semantic_md.models import FailureInfo, SheetModel
from excel_semantic_md.render.types import RenderSheetResult


class GitHubCopilotSdkAdapter:
    """Run one Copilot SDK session per sheet."""

    async def run_sheet_async(
        self,
        sheet: SheetModel,
        render_result: RenderSheetResult | None,
        *,
        options: LlmRunOptions | None = None,
    ) -> LlmRunResult:
        run_options = options or LlmRunOptions()
        attachments = build_llm_attachments(
            sheet,
            render_result,
            max_images_per_sheet=run_options.max_images_per_sheet,
        )
        llm_input = build_llm_input(sheet, attachments)
        prompt = build_sheet_prompt(llm_input)
        result: LlmRunResult | None = None
        client: Any | None = None

        try:
            copilot_client_type, permission_handler = _import_copilot_sdk()
            client = copilot_client_type()
            await client.start()
        except Exception as exc:
            return LlmRunResult(
                status="failed",
                attempts=1,
                failure=FailureInfo(
                    stage="llm",
                    message="Failed to initialize GitHub Copilot SDK client.",
                    details={"sheet_name": sheet.name, "error": str(exc)},
                ),
            )

        attempts = 0
        used_model: str | None = None
        try:
            session = await client.create_session(**_build_session_kwargs(run_options, permission_handler))
            used_model = await _get_used_model(session)
            attachment_payload = [_attachment_payload(item) for item in attachments]

            last_error: ValueError | None = None
            for _ in range(2):
                attempts += 1
                response = await _send_and_wait(session, prompt, attachment_payload)
                response_text = _extract_response_text(response)
                try:
                    parsed = parse_llm_response(response_text)
                    result = LlmRunResult(status="succeeded", attempts=attempts, response=parsed, used_model=used_model)
                    break
                except ValueError as exc:
                    last_error = exc

            if result is None:
                assert last_error is not None
                result = LlmRunResult(
                    status="failed",
                    attempts=attempts,
                    failure=FailureInfo(
                        stage="llm",
                        message="LLM response validation failed after retry.",
                        details={"sheet_name": sheet.name, "error": str(last_error)},
                    ),
                    used_model=used_model,
                )
        except Exception as exc:
            result = LlmRunResult(
                status="failed",
                attempts=max(attempts, 1),
                failure=FailureInfo(
                    stage="llm",
                    message="GitHub Copilot SDK execution failed.",
                    details={"sheet_name": sheet.name, "error": str(exc)},
                ),
                used_model=used_model,
            )
        finally:
            if client is not None:
                stop_result = await _stop_client(client, result, sheet_name=sheet.name, attempts=max(attempts, 1))
                if stop_result is not None:
                    result = stop_result

        assert result is not None
        return result

    def run_sheet(
        self,
        sheet: SheetModel,
        render_result: RenderSheetResult | None,
        *,
        options: LlmRunOptions | None = None,
    ) -> LlmRunResult:
        return asyncio.run(self.run_sheet_async(sheet, render_result, options=options))


def _import_copilot_sdk() -> tuple[type[Any], Any]:
    from copilot import CopilotClient
    from copilot.session import PermissionHandler

    return CopilotClient, PermissionHandler


def _build_session_kwargs(options: LlmRunOptions, permission_handler: Any) -> dict[str, Any]:
    session_kwargs: dict[str, Any] = {}
    if options.model is not None:
        session_kwargs["model"] = options.model
    if options.vision_model is not None:
        session_kwargs["vision_model"] = options.vision_model
    return session_kwargs


async def _send_and_wait(session: Any, prompt: str, attachments: list[dict[str, str]]) -> Any:
    if attachments:
        try:
            return await session.send_and_wait(prompt, attachments=attachments)
        except TypeError:
            return await session.send_and_wait({"prompt": prompt, "attachments": attachments})
    return await session.send_and_wait(prompt)


def _attachment_payload(attachment: Any) -> dict[str, str]:
    return {
        "type": "file",
        "path": attachment.path,
    }


async def _get_used_model(session: Any) -> str | None:
    try:
        rpc = getattr(session, "rpc", None)
        if rpc is None or getattr(rpc, "model", None) is None:
            return None
        current = await rpc.model.get_current()
    except Exception:
        return None
    return getattr(current, "model_id", None)


def _extract_response_text(response: Any) -> str:
    if response is None:
        raise ValueError("LLM response did not include an assistant message")

    data = getattr(response, "data", response)
    content = getattr(data, "content", None)
    if isinstance(content, str):
        return content
    if isinstance(response, str):
        return response
    raise ValueError("LLM response did not include assistant message content")


async def _stop_client(
    client: Any,
    result: LlmRunResult | None,
    *,
    sheet_name: str,
    attempts: int,
) -> LlmRunResult | None:
    try:
        await client.stop()
    except Exception as exc:
        if result is not None and result.failure is not None:
            failure = FailureInfo(
                stage=result.failure.stage,
                message=result.failure.message,
                details={**result.failure.details, "cleanup_error": str(exc)},
            )
            return LlmRunResult(
                status="failed",
                attempts=result.attempts,
                response=result.response,
                failure=failure,
                used_model=result.used_model,
            )
        return LlmRunResult(
            status="failed",
            attempts=attempts,
            response=None if result is None else result.response,
            failure=FailureInfo(
                stage="llm_cleanup",
                message="GitHub Copilot SDK cleanup failed.",
                details={"sheet_name": sheet_name, "cleanup_error": str(exc)},
            ),
            used_model=None if result is None else result.used_model,
        )
    return None
