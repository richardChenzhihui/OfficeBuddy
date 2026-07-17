"""Schemas for harness-intercepted tools (plan / ask_user / render_preview /
report_verification). Their handlers live in the agent loop, which owns UI and
render state — they are NOT dispatched through the registry."""
from typing import List, Literal, Optional

from pydantic import BaseModel, Field


class ProposePlanInput(BaseModel):
    steps: List[str] = Field(
        ...,
        description=(
            "Concrete, verifiable step descriptions in execution order, e.g. "
            "['Replace paragraph 0 text with the new title', "
            "'Set heading font to Times New Roman 14pt bold']"
        ),
    )


class UpdatePlanInput(BaseModel):
    step_index: int = Field(..., description="0-based index of the step")
    status: Literal["in_progress", "done", "blocked"] = Field(
        ...,
        description=(
            "Set 'in_progress' when starting a step. Set 'done' when you believe "
            "the step is complete — the harness will render and visually verify "
            "it; if verification fails the result tells you exactly what to fix "
            "and the step stays in_progress."
        ),
    )


class AskUserOption(BaseModel):
    label: str = Field(..., description="Short option label")
    description: str = Field("", description="What choosing this option means")
    is_default_safe: bool = Field(
        False,
        description=(
            "True if this option is a safe default that can be auto-selected in "
            "non-interactive mode"
        ),
    )


class AskUserInput(BaseModel):
    question: str = Field(..., description="The question for the user, concise and specific")
    kind: Literal["free_form", "multiple_choice"] = Field(
        "multiple_choice", description="Question type"
    )
    options: Optional[List[AskUserOption]] = Field(
        None, description="2-4 options for multiple_choice"
    )
    allow_multiple: bool = Field(False, description="Allow selecting several options")


class RenderPreviewInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    pages: Optional[List[int]] = Field(
        None,
        description="0-based page indices to view (max 3). Omit for the first pages.",
    )
    sheet: Optional[str] = Field(
        None, description="Excel only: sheet to render (defaults to active sheet)"
    )


class VerificationProblem(BaseModel):
    page: int = Field(..., description="0-based page index where the problem is")
    element_hint: str = Field(..., description="Which element, e.g. 'first heading paragraph'")
    description: str = Field(..., description="Concrete discrepancy, e.g. 'font is Calibri, expected Times New Roman'")
    severity: Literal["blocking", "minor"] = Field(..., description="blocking = must repair")


class ReportVerificationInput(BaseModel):
    passed: bool = Field(..., description="True only if the intended change is correctly visible")
    problems: List[VerificationProblem] = Field(
        default_factory=list, description="All discrepancies found (empty if passed)"
    )
    confidence: float = Field(..., description="0.0-1.0 confidence in this verdict")


def interaction_tool_schemas() -> List[dict]:
    """Anthropic tool definitions for the harness-intercepted tools."""
    return [
        {
            "name": "propose_plan",
            "description": (
                "Propose a step-by-step plan BEFORE editing on any multi-step "
                "task. Steps are shown to the user as a live checklist."
            ),
            "input_schema": ProposePlanInput.model_json_schema(),
        },
        {
            "name": "update_plan",
            "description": (
                "Update a plan step's status. Setting 'done' triggers rendering "
                "and visual verification of your work — the result tells you "
                "whether it really passed."
            ),
            "input_schema": UpdatePlanInput.model_json_schema(),
        },
        {
            "name": "ask_user",
            "description": (
                "Ask the user a clarifying question or request a decision. Use "
                "BEFORE guessing on ambiguous instructions, and before "
                "destructive actions. Prefer multiple_choice with 2-4 concrete "
                "options."
            ),
            "input_schema": AskUserInput.model_json_schema(),
        },
        {
            "name": "render_preview",
            "description": (
                "Look at the document as the real Word/Excel app renders it "
                "(returns page screenshots). Use when you need to see current "
                "visual state mid-task. Max 3 pages per call."
            ),
            "input_schema": RenderPreviewInput.model_json_schema(),
        },
    ]


REPORT_VERIFICATION_TOOL = {
    "name": "report_verification",
    "description": "Report the verification verdict for the edit under review.",
    "input_schema": ReportVerificationInput.model_json_schema(),
}
