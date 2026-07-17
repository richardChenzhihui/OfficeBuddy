"""Scripted fake LLM matching the surface AgentSession calls."""
import itertools
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class FakeBlock:
    type: str
    text: Optional[str] = None
    id: Optional[str] = None
    name: Optional[str] = None
    input: Optional[Dict[str, Any]] = None


_ids = itertools.count(1)


def text_block(text: str) -> FakeBlock:
    return FakeBlock(type="text", text=text)


def tool_use(name: str, tool_input: Dict[str, Any]) -> FakeBlock:
    return FakeBlock(type="tool_use", id=f"tu_{next(_ids)}", name=name, input=tool_input)


@dataclass
class FakeMessage:
    content: List[FakeBlock]
    stop_reason: str = "end_turn"


@dataclass
class FakeLLM:
    """Pops scripted responses. Verifier calls (forced report_verification)
    are routed to verify_script so main-loop scripts stay readable."""

    script: List[FakeMessage] = field(default_factory=list)
    verify_script: List[FakeMessage] = field(default_factory=list)
    calls: List[Dict[str, Any]] = field(default_factory=list)

    def create(self, system, messages, tools=None, tool_choice=None, max_tokens=None):
        self.calls.append(
            {
                "system": system,
                "messages": [dict(m) for m in messages],
                "tools": [t["name"] for t in (tools or [])],
                "tool_choice": tool_choice,
            }
        )
        if tool_choice and tool_choice.get("name") == "report_verification":
            if not self.verify_script:
                raise AssertionError("FakeLLM: unexpected verifier call")
            return self.verify_script.pop(0)
        if not self.script:
            raise AssertionError("FakeLLM: script exhausted")
        return self.script.pop(0)


def verify_verdict(passed: bool, problems=None, confidence=0.9) -> FakeMessage:
    return FakeMessage(
        content=[
            tool_use(
                "report_verification",
                {"passed": passed, "problems": problems or [], "confidence": confidence},
            )
        ],
        stop_reason="tool_use",
    )
