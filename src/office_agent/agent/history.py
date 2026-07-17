"""Conversation history management for the Anthropic Messages format.

Invariants:
- The assistant turn's FULL content (text + tool_use + thinking) is appended
  as returned.
- ALL tool_result blocks answering one assistant turn go into a SINGLE user
  message.
- Images never enter the main history except inside render_preview results,
  and old tool_result bodies are elided when history grows large.
"""
from typing import Any, Dict, List

from ..llm.client import content_block_to_dict

# Rough char budget before old tool results get elided (~1M-token model, but
# keep sessions snappy and cheap).
DEFAULT_CHAR_BUDGET = 400_000
KEEP_RECENT_MESSAGES = 12


class History:
    def __init__(self, char_budget: int = DEFAULT_CHAR_BUDGET):
        self.messages: List[Dict[str, Any]] = []
        self.char_budget = char_budget

    def add_user_text(self, text: str) -> None:
        self.messages.append({"role": "user", "content": [{"type": "text", "text": text}]})

    def add_assistant(self, content_blocks: List[Any]) -> None:
        self.messages.append(
            {
                "role": "assistant",
                "content": [content_block_to_dict(b) for b in content_blocks],
            }
        )

    def add_tool_results(self, results: List[Dict[str, Any]]) -> None:
        """results: [{'tool_use_id', 'content', 'is_error'?}] — one user msg."""
        blocks = []
        for r in results:
            block: Dict[str, Any] = {
                "type": "tool_result",
                "tool_use_id": r["tool_use_id"],
                "content": r["content"],
            }
            if r.get("is_error"):
                block["is_error"] = True
            blocks.append(block)
        self.messages.append({"role": "user", "content": blocks})
        self.trim_if_needed()

    def _size(self) -> int:
        import json

        return len(json.dumps(self.messages, default=str))

    def trim_if_needed(self) -> None:
        if self._size() <= self.char_budget:
            return
        # Elide old tool_result bodies (never the most recent turns).
        for msg in self.messages[:-KEEP_RECENT_MESSAGES]:
            if msg["role"] != "user":
                continue
            for block in msg["content"]:
                if isinstance(block, dict) and block.get("type") == "tool_result":
                    content = block.get("content")
                    if isinstance(content, list) or (
                        isinstance(content, str) and len(content) > 200
                    ):
                        block["content"] = "[elided earlier tool output]"
            if self._size() <= self.char_budget:
                break
