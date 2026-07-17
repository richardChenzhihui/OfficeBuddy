"""LLM client wrapper over the Anthropic SDK pointed at MiniMax's
Anthropic-compatible endpoint. Only GA Messages API surface is used
(system / messages / tools / tool_choice / image blocks) so any faithful
compatibility layer works.
"""
import time
from typing import Any, Dict, List, Optional

import anthropic

from ..config import Config

TRANSIENT = (
    anthropic.RateLimitError,
    anthropic.InternalServerError,
    anthropic.APIConnectionError,
    anthropic.APITimeoutError,
)


class LLM:
    """Thin wrapper the agent loop calls; FakeLLM in tests mimics this surface."""

    def __init__(self, config: Config):
        config.require_api_key()
        self.config = config
        self.client = anthropic.Anthropic(
            api_key=config.api_key, base_url=config.base_url
        )

    def create(
        self,
        system: str,
        messages: List[Dict[str, Any]],
        tools: Optional[List[Dict[str, Any]]] = None,
        tool_choice: Optional[Dict[str, Any]] = None,
        max_tokens: Optional[int] = None,
    ):
        kwargs: Dict[str, Any] = {
            "model": self.config.model,
            "max_tokens": max_tokens or self.config.max_tokens,
            "system": system,
            "messages": messages,
        }
        if tools:
            kwargs["tools"] = tools
        if tool_choice:
            kwargs["tool_choice"] = tool_choice

        delay = 2.0
        for attempt in range(4):
            try:
                return self.client.messages.create(**kwargs)
            except TRANSIENT as exc:
                if attempt == 3:
                    raise
                time.sleep(delay)
                delay *= 2


def content_block_to_dict(block: Any) -> Dict[str, Any]:
    """Normalize SDK content blocks (or already-plain dicts) for re-sending."""
    if isinstance(block, dict):
        return block
    btype = getattr(block, "type", None)
    if btype == "text":
        return {"type": "text", "text": block.text}
    if btype == "tool_use":
        return {"type": "tool_use", "id": block.id, "name": block.name, "input": block.input}
    if btype == "thinking":
        result = {"type": "thinking", "thinking": block.thinking}
        signature = getattr(block, "signature", None)
        if signature:
            result["signature"] = signature
        return result
    # Unknown block type: best-effort passthrough
    dump = getattr(block, "model_dump", None)
    return dump() if dump else {"type": str(btype)}
