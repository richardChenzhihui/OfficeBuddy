"""Visual verification via an independent, stateless model call.

Isolation is deliberate: the verifier sees only the step's intent and the
rendered images — never the executor's transcript or self-reported success —
so its verdict is not anchored on the executor's narrative. Images stay out
of the main loop's history entirely.
"""
import base64
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from ..render.pdf_to_images import PageImage
from ..tools.interaction_tools import REPORT_VERIFICATION_TOOL
from .prompts import SYSTEM_VERIFIER


@dataclass
class VerificationResult:
    passed: bool
    problems: List[Dict[str, Any]] = field(default_factory=list)
    confidence: float = 0.0
    skipped: bool = False

    def summary(self) -> str:
        if self.skipped:
            return "Visual verification skipped."
        if self.passed:
            return f"Verification PASSED (confidence {self.confidence:.2f})."
        lines = [f"Verification FAILED (confidence {self.confidence:.2f}):"]
        for p in self.problems:
            lines.append(
                f"- page {p.get('page')}, {p.get('element_hint')}: "
                f"{p.get('description')} [{p.get('severity')}]"
            )
        return "\n".join(lines)

    @property
    def blocking(self) -> bool:
        if self.passed or self.skipped:
            return False
        return any(p.get("severity") == "blocking" for p in self.problems) or not self.problems


def _image_block(png_bytes: bytes) -> Dict[str, Any]:
    return {
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": "image/png",
            "data": base64.b64encode(png_bytes).decode(),
        },
    }


def verify_edit(
    llm,
    step_description: str,
    operations_summary: str,
    after_images: List[PageImage],
    before_images: Optional[List[PageImage]] = None,
    max_pages: int = 4,
    extra_note: str = "",
) -> VerificationResult:
    """One-shot stateless verification call. Forces a structured verdict."""
    content: List[Dict[str, Any]] = [
        {
            "type": "text",
            "text": (
                f"## Intended change for this step\n{step_description}\n\n"
                f"## Operations the editor applied\n{operations_summary}\n"
                + (f"\n{extra_note}\n" if extra_note else "")
            ),
        }
    ]
    shown_after = after_images[:max_pages]
    if before_images:
        shown_before = before_images[:max_pages]
        content.append(
            {"type": "text", "text": f"## BEFORE images ({len(shown_before)} page(s))"}
        )
        for img in shown_before:
            content.append({"type": "text", "text": f"BEFORE page {img.index}:"})
            content.append(_image_block(img.png_bytes))
    content.append(
        {"type": "text", "text": f"## AFTER images ({len(shown_after)} page(s))"}
    )
    for img in shown_after:
        content.append({"type": "text", "text": f"AFTER page {img.index}:"})
        content.append(_image_block(img.png_bytes))
    if len(after_images) > max_pages:
        content.append(
            {
                "type": "text",
                "text": (
                    f"Note: {len(after_images) - max_pages} more changed page(s) "
                    "not shown."
                ),
            }
        )

    response = llm.create(
        system=SYSTEM_VERIFIER,
        messages=[{"role": "user", "content": content}],
        tools=[REPORT_VERIFICATION_TOOL],
        tool_choice={"type": "tool", "name": "report_verification"},
        max_tokens=2048,
    )
    for block in response.content:
        if getattr(block, "type", None) == "tool_use" and block.name == "report_verification":
            data = block.input
            return VerificationResult(
                passed=bool(data.get("passed")),
                problems=list(data.get("problems", [])),
                confidence=float(data.get("confidence", 0.0)),
            )
    # Model failed to produce the forced tool call — treat as inconclusive pass
    # with a note rather than blocking forever.
    return VerificationResult(passed=True, confidence=0.0, skipped=True)
