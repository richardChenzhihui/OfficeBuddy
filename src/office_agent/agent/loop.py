"""The agent harness loop.

Control flow per user turn:
  intake -> (model plans via propose_plan / clarifies via ask_user)
         -> execute tool calls
         -> update_plan(done) triggers render + independent visual verification
         -> failures come back as precise, actionable tool results
         -> escalation ladder prevents brute-force retries
         -> end_turn safety net verifies any unverified visual changes

Harness-intercepted tools (propose_plan / update_plan / ask_user /
render_preview) are handled here because they need UI and render state.
"""
import json
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from ..config import Config
from ..llm.client import LLM
from ..render import Renderer, diff_pages, highlight_region
from ..render.applescript import RenderError
from ..render.pdf_to_images import PageImage
from ..tools import REGISTRY, ToolContext
from ..tools.interaction_tools import (
    AskUserInput,
    ProposePlanInput,
    RenderPreviewInput,
    UpdatePlanInput,
    interaction_tool_schemas,
)
from .budget import Action, BudgetTracker, normalize_signature
from .history import History
from .plan import Plan
from .prompts import SYSTEM_EXECUTOR
from .verifier import VerificationResult, verify_edit

# Mutations whose effect is purely data (verified by tool results / re-reads);
# everything else gets visual verification.
NON_VISUAL_TOOLS = {"excel_write_cells", "excel_edit_formula"}

HARNESS_TOOLS = {"propose_plan", "update_plan", "ask_user", "render_preview"}


@dataclass
class TurnResult:
    text: str = ""
    plan: Optional[Plan] = None
    saved_paths: List[str] = field(default_factory=list)
    aborted: bool = False


class BaseUI:
    """Minimal UI protocol; CLI and headless implementations override."""

    def emit_text(self, text: str) -> None: ...

    def notify(self, message: str) -> None: ...

    def tool_call(self, name: str, tool_input: Dict[str, Any]) -> None: ...

    def tool_result(self, name: str, result: Dict[str, Any]) -> None: ...

    def plan_update(self, plan: Plan) -> None: ...

    def ask_user(self, params: AskUserInput) -> Dict[str, Any]:
        """Return {'answer': str | list}. Headless default: safe option or note."""
        if params.options:
            safe = [o for o in params.options if o.is_default_safe]
            if safe:
                return {
                    "answer": safe[0].label,
                    "note": "auto-selected safe default (non-interactive)",
                }
        return {
            "answer": "",
            "note": (
                "Non-interactive session: no user available. Proceed with your "
                "best judgement and clearly flag the assumption in your summary."
            ),
        }


class AgentSession:
    def __init__(
        self,
        config: Config,
        ui: Optional[BaseUI] = None,
        llm: Optional[Any] = None,
        ctx: Optional[ToolContext] = None,
    ):
        self.config = config
        self.ui = ui or BaseUI()
        self.llm = llm or LLM(config)
        self.ctx = ctx or ToolContext()
        self.history = History()
        self.plan: Optional[Plan] = None
        self.budget = BudgetTracker(config.budgets)
        self.renderers: Dict[str, Renderer] = {}
        self.visual_dirty: Dict[str, set] = {}
        self.pending_ops: Dict[str, List[str]] = {}
        self.saved_paths: List[str] = []
        self.abort_requested = False

    # ------------------------------------------------------------------ turn

    def run_turn(self, instruction: str, file_path: Optional[str] = None) -> TurnResult:
        text = instruction
        if file_path:
            text += f"\n\n[Document to edit: {file_path}]"
        self.history.add_user_text(text)
        self.saved_paths = []
        final_text = ""

        while True:
            response = self.llm.create(
                system=SYSTEM_EXECUTOR,
                messages=self.history.messages,
                tools=self._tools(),
            )
            self.history.add_assistant(list(response.content))

            for block in response.content:
                if getattr(block, "type", None) == "text" and block.text.strip():
                    final_text = block.text
                    self.ui.emit_text(block.text)

            tool_uses = [
                b for b in response.content if getattr(b, "type", None) == "tool_use"
            ]
            if not tool_uses:
                repair = self._end_turn_safety_net()
                if repair and not self.budget.end_turn_exhausted():
                    self.budget.record_end_turn_repair()
                    self.history.add_user_text(repair)
                    continue
                if repair:
                    self.ui.notify(
                        "验证仍未通过且修复次数已达上限，停止本回合。"
                        "文档保持当前状态（原文件未动，可用 undo 回退）。"
                    )
                    return TurnResult(
                        text=final_text,
                        plan=self.plan,
                        saved_paths=self.saved_paths,
                        aborted=True,
                    )
                return TurnResult(
                    text=final_text, plan=self.plan, saved_paths=self.saved_paths
                )

            results = []
            for tu in tool_uses:
                content, is_error = self._handle_tool(tu.name, dict(tu.input))
                results.append(
                    {"tool_use_id": tu.id, "content": content, "is_error": is_error}
                )
            self.history.add_tool_results(results)

            if self.abort_requested:
                self.ui.notify("按用户要求中止任务。")
                return TurnResult(
                    text=final_text,
                    plan=self.plan,
                    saved_paths=self.saved_paths,
                    aborted=True,
                )

    # ----------------------------------------------------------- tool routing

    def _tools(self) -> List[Dict[str, Any]]:
        return REGISTRY.to_anthropic_tools() + interaction_tool_schemas()

    def _current_step(self) -> int:
        return self.plan.current_index if self.plan else 0

    def _handle_tool(self, name: str, tool_input: Dict[str, Any]):
        """Returns (content, is_error) for the tool_result block.

        Never raises: any unexpected exception becomes an error envelope, so
        every tool_use always gets its matching tool_result.
        """
        try:
            return self._handle_tool_inner(name, tool_input)
        except Exception as exc:  # noqa: BLE001 — pairing invariant beats purity
            return (
                json.dumps(
                    {
                        "success": False,
                        "error": f"Internal harness error in {name}: {exc}",
                        "error_type": type(exc).__name__,
                    },
                    ensure_ascii=False,
                ),
                True,
            )

    def _handle_tool_inner(self, name: str, tool_input: Dict[str, Any]):
        self.ui.tool_call(name, tool_input)

        if self.budget.task_exhausted() and name not in HARNESS_TOOLS:
            return (
                json.dumps(
                    {
                        "success": False,
                        "error": (
                            f"Task tool-call budget "
                            f"({self.config.budgets.max_tool_calls_per_task}) "
                            "exhausted. Stop editing: summarize progress, state "
                            "what remains, and ask the user how to proceed."
                        ),
                    }
                ),
                True,
            )

        if name == "propose_plan":
            return self._on_propose_plan(tool_input)
        if name == "update_plan":
            return self._on_update_plan(tool_input)
        if name == "ask_user":
            return self._on_ask_user(tool_input)
        if name == "render_preview":
            return self._on_render_preview(tool_input)

        # Save gate: intercept BEFORE dispatch whenever the target is an
        # existing file this session hasn't already written.
        if name == "save_document":
            gate = self._save_gate(tool_input)
            if gate is not None:
                return gate

        self.budget.record_tool_call(self._current_step())
        result = REGISTRY.dispatch(self.ctx, name, tool_input)
        self.ui.tool_result(name, result)

        tool_def = REGISTRY.tools.get(name)
        if result.get("success") and tool_def and tool_def.mutates:
            doc_id = tool_input.get("doc_id", "")
            self.visual_dirty.setdefault(doc_id, set()).add(name)
            self.pending_ops.setdefault(doc_id, []).append(
                f"{name}({json.dumps(tool_input, ensure_ascii=False, default=str)[:300]})"
            )
        if result.get("success") and name == "save_document":
            self.saved_paths.append(result.get("saved_path", ""))

        if not result.get("success"):
            signature = normalize_signature(name, str(result.get("error", "")))
            action = self.budget.record_failure(self._current_step(), signature)
            result = self._apply_escalation(result, action, signature)
            return json.dumps(result, ensure_ascii=False, default=str), True

        return json.dumps(result, ensure_ascii=False, default=str), False

    # ------------------------------------------------------ harness handlers

    def _on_propose_plan(self, tool_input):
        try:
            params = ProposePlanInput(**tool_input)
        except Exception as exc:
            return json.dumps({"success": False, "error": str(exc)}), True
        self.plan = Plan.from_descriptions(params.steps)
        # Fresh plan -> fresh per-step budgets (task-level counters persist).
        self.budget.reset_steps()
        self.ui.plan_update(self.plan)
        return (
            json.dumps({"success": True, "steps_recorded": len(params.steps)}),
            False,
        )

    def _on_update_plan(self, tool_input):
        try:
            params = UpdatePlanInput(**tool_input)
        except Exception as exc:
            return json.dumps({"success": False, "error": str(exc)}), True
        if self.plan is None:
            return (
                json.dumps(
                    {"success": False, "error": "No plan exists: call propose_plan first."}
                ),
                True,
            )
        try:
            self.plan.set_status(params.step_index, params.status)
        except ValueError as exc:
            return json.dumps({"success": False, "error": str(exc)}), True

        if params.status != "done":
            self.ui.plan_update(self.plan)
            return json.dumps({"success": True}), False

        # 'done' is a verification checkpoint.
        step_desc = self.plan.steps[params.step_index].description
        try:
            verdicts = self._verify_pending(step_desc)
        except Exception as exc:  # keep tool_use/tool_result pairing intact
            self.plan.set_status(params.step_index, "in_progress")
            self.ui.plan_update(self.plan)
            return (
                json.dumps(
                    {
                        "success": False,
                        "error": (
                            f"Verification errored ({type(exc).__name__}: {exc}). "
                            "The step remains in_progress; check the document "
                            "state (render_preview / re-read) before retrying."
                        ),
                    },
                    ensure_ascii=False,
                ),
                True,
            )
        blocking = [v for v in verdicts.values() if v.blocking]
        if blocking:
            self.plan.set_status(params.step_index, "in_progress")
            self.ui.plan_update(self.plan)
            detail = "\n\n".join(v.summary() for v in verdicts.values())
            signature = normalize_signature("verify", detail)
            action = self.budget.record_failure(params.step_index, signature)
            result = {
                "success": False,
                "verification": detail,
                "error": (
                    "Step NOT done: visual verification failed. Fix the specific "
                    "problems listed, then set the step to done again."
                ),
            }
            result = self._apply_escalation(result, action, signature)
            return json.dumps(result, ensure_ascii=False), True

        self.ui.plan_update(self.plan)
        summaries = [v.summary() for v in verdicts.values()] or [
            "No visual changes pending verification."
        ]
        return (
            json.dumps({"success": True, "verification": "\n".join(summaries)}),
            False,
        )

    def _on_ask_user(self, tool_input):
        try:
            params = AskUserInput(**tool_input)
        except Exception as exc:
            return json.dumps({"success": False, "error": str(exc)}), True
        answer = self.ui.ask_user(params)
        return json.dumps({"success": True, **answer}, ensure_ascii=False), False

    def _on_render_preview(self, tool_input):
        try:
            params = RenderPreviewInput(**tool_input)
        except Exception as exc:
            return json.dumps({"success": False, "error": str(exc)}), True
        try:
            # Whitelist the sheet name BEFORE it can reach the AppleScript
            # template: it is a free-form, model-chosen string.
            if params.sheet:
                session = self.ctx.sessions.get(params.doc_id)
                if (
                    session.doc_type != "excel"
                    or params.sheet not in session.doc.sheetnames
                ):
                    return (
                        json.dumps(
                            {
                                "success": False,
                                "error": (
                                    f"Sheet '{params.sheet}' not found: workbook "
                                    f"sheets are {getattr(session.doc, 'sheetnames', [])}."
                                ),
                            },
                            ensure_ascii=False,
                        ),
                        True,
                    )
            renderer = self._renderer(params.doc_id)
            images = renderer.render(
                sheet=params.sheet, timeout=self.config.render_timeout
            )
        except (RenderError, KeyError) as exc:
            return json.dumps({"success": False, "error": str(exc)}), True
        wanted = params.pages if params.pages else [i.index for i in images[:3]]
        selected = [img for img in images if img.index in set(wanted)][:3]
        content: List[Dict[str, Any]] = [
            {
                "type": "text",
                "text": (
                    f"Rendered {len(images)} page(s); showing "
                    f"{[i.index for i in selected]}."
                ),
            }
        ]
        import base64

        for img in selected:
            content.append({"type": "text", "text": f"Page {img.index}:"})
            content.append(
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": base64.b64encode(img.png_bytes).decode(),
                    },
                }
            )
        return content, False

    # ----------------------------------------------------------- escalation

    def _apply_escalation(self, result: Dict[str, Any], action: Action, signature: str):
        if action == Action.SWITCH_STRATEGY:
            result["escalation"] = (
                f"You have now failed twice the same way ({signature}). Do NOT "
                "repeat the same approach or selector. Alternatives: use a "
                "text_match selector instead of indices (or vice versa); re-run "
                "get_structure to re-ground yourself; try a smaller, more "
                "targeted edit; or use render_preview to look at the document."
            )
        elif action == Action.ASK_USER:
            step_budget = self.budget.step(self._current_step())
            step_budget.user_asks += 1
            if step_budget.user_asks > 2:
                # The user has already been asked twice about this step:
                # stop instead of nagging forever.
                self.abort_requested = True
                result["escalation"] = (
                    "Repeated failures persist after user guidance. The task "
                    "is being aborted; summarize progress and what remains."
                )
                return result
            answer = self.ui.ask_user(
                AskUserInput(
                    question=(
                        "The agent repeatedly failed at this step "
                        f"({signature}). How should it proceed?"
                    ),
                    kind="multiple_choice",
                    options=[
                        {
                            "label": "Try a different approach",
                            "description": "One more attempt with a changed strategy",
                            "is_default_safe": True,
                        },
                        {
                            "label": "Skip this step",
                            "description": "Continue with the rest of the plan",
                            "is_default_safe": False,
                        },
                        {
                            "label": "Abort the task",
                            "description": "Stop here and report progress",
                            "is_default_safe": False,
                        },
                    ],
                    allow_multiple=False,
                )
            )
            answer_text = str(answer.get("answer", "")).lower()
            if answer_text.startswith("abort"):
                # Enforce the decision in code — don't rely on the model.
                self.abort_requested = True
            elif answer_text.startswith("skip") and self.plan is not None:
                self.plan.set_status(self._current_step(), "blocked")
                self.ui.plan_update(self.plan)
            result["escalation"] = (
                f"Repeated failures. The user was asked and answered: "
                f"{json.dumps(answer, ensure_ascii=False)}. Follow this decision."
            )
        return result

    # ---------------------------------------------------------- verification

    def _renderer(self, doc_id: str) -> Renderer:
        if doc_id not in self.renderers:
            session = self.ctx.sessions.get(doc_id)
            self.renderers[doc_id] = Renderer(session)
        return self.renderers[doc_id]

    def _docs_needing_verify(self) -> List[str]:
        return [d for d, tools in self.visual_dirty.items() if tools]

    def _verify_pending(self, step_description: str) -> Dict[str, VerificationResult]:
        verdicts: Dict[str, VerificationResult] = {}
        for doc_id in self._docs_needing_verify():
            verdicts[doc_id] = self._verify_doc(doc_id, step_description)
        return verdicts

    def _verify_doc(self, doc_id: str, step_description: str) -> VerificationResult:
        tools_used = self.visual_dirty.get(doc_id, set())
        ops = self.pending_ops.get(doc_id, [])

        if not self.config.visual_verify or tools_used <= NON_VISUAL_TOOLS:
            # Pure data edits: tool-result verification is sufficient.
            self.visual_dirty[doc_id] = set()
            self.pending_ops[doc_id] = []
            return VerificationResult(passed=True, confidence=1.0, skipped=True)

        self.ui.notify("Rendering document for visual verification…")
        renderer = self._renderer(doc_id)
        try:
            after = renderer.render(timeout=self.config.render_timeout)
        except RenderError as exc:
            return VerificationResult(
                passed=False,
                problems=[
                    {
                        "page": 0,
                        "element_hint": "document",
                        "description": f"Rendering failed: {exc}",
                        "severity": "blocking",
                    }
                ],
                confidence=1.0,
            )

        # Diff against the last VERIFIED baseline, not merely the previous
        # render — ad hoc render_preview calls must not shrink the diff.
        before = renderer.baseline
        extra_note = ""
        if before is not None:
            diffs = diff_pages(before, after)
            changed = [d for d in diffs if d.changed]
            if changed:
                changed_idx = {d.page_index for d in changed}
                after_sel = [i for i in after if i.index in changed_idx]
                before_sel = [i for i in before if i.index in changed_idx]
                annotated = []
                for img in after_sel:
                    diff = next(d for d in changed if d.page_index == img.index)
                    if diff.bbox:
                        out = img.path.with_name(f"annotated_{img.path.name}")
                        highlight_region(img, diff.bbox, out)
                        annotated.append(
                            PageImage(img.index, out, img.width, img.height)
                        )
                    else:
                        annotated.append(img)
                after_images, before_images = annotated, before_sel
            else:
                after_images, before_images = after[: self.config.max_verify_pages], None
                extra_note = (
                    "NOTE: pixel diff detected NO visible change since the last "
                    "verified render, despite edits being applied. Check "
                    "carefully whether the intended change is actually present."
                )
        else:
            after_images, before_images = after[: self.config.max_verify_pages], None

        self.ui.notify(
            f"Verifying {len(after_images)} page(s) with independent reviewer…"
        )
        verdict = verify_edit(
            self.llm,
            step_description,
            "\n".join(ops) or "(no operation summary)",
            after_images,
            before_images,
            max_pages=self.config.max_verify_pages,
            extra_note=extra_note,
        )
        self.ui.notify(verdict.summary())
        if verdict.passed:
            self.visual_dirty[doc_id] = set()
            self.pending_ops[doc_id] = []
            renderer.baseline = after
        return verdict

    def _end_turn_safety_net(self) -> Optional[str]:
        """If the model ends the turn with unverified visual edits, verify now."""
        pending = self._docs_needing_verify()
        if not pending or not self.config.visual_verify:
            return None
        step_desc = "Final state should match the user's instruction."
        verdicts = self._verify_pending(step_desc)
        blocking = {d: v for d, v in verdicts.items() if v.blocking}
        if not blocking:
            return None
        detail = "\n\n".join(v.summary() for v in blocking.values())
        signature = normalize_signature("verify", detail)
        action = self.budget.record_failure(self._current_step(), signature)
        message = (
            "[harness] You ended the turn, but visual verification of your edits "
            f"FAILED:\n{detail}\nFix the specific problems, then finish."
        )
        escalated = self._apply_escalation({}, action, signature)
        if "escalation" in escalated:
            message += "\n" + escalated["escalation"]
        return message

    # -------------------------------------------------------------- save gate

    def _save_gate(self, tool_input: Dict[str, Any]):
        """Gate BEFORE dispatch for any save whose target is an existing file
        this session hasn't already written. Returns None to allow dispatch."""
        from pathlib import Path

        doc_id = tool_input.get("doc_id", "")
        session = self.ctx.sessions.sessions.get(doc_id)
        if session is None:
            return None  # dispatch will produce the unknown-doc error
        raw = tool_input.get("path")
        target = (
            Path(raw).expanduser().resolve() if raw else session.default_output_path()
        )
        is_original = target == session.original_path
        if not target.exists() or str(target) in session.written_paths:
            return None  # new file, or one we wrote earlier this session
        if not tool_input.get("overwrite"):
            return None  # session.save_to will refuse with an actionable error

        # --yes only pre-approves overwriting the document's own original file,
        # never arbitrary other existing files.
        if is_original and self.config.auto_approve_overwrite:
            return None
        what = (
            f"原文件 {target}" if is_original else f"已存在的其他文件 {target}"
        )
        if self.config.non_interactive:
            return (
                json.dumps(
                    {
                        "success": False,
                        "error": (
                            f"Overwriting {target} requires "
                            + ("--yes" if is_original else "interactive user approval")
                            + " in non-interactive mode. Save to a new path "
                            "instead (omit 'path' to write <name>.edited.<ext>)."
                        ),
                    }
                ),
                True,
            )
        answer = self.ui.ask_user(
            AskUserInput(
                question=f"Agent 想要覆盖{what}，允许吗？",
                kind="multiple_choice",
                options=[
                    {
                        "label": "另存为新文件",
                        "description": "写入 <name>.edited.<ext>，目标文件不动",
                        "is_default_safe": True,
                    },
                    {
                        "label": "Yes, overwrite",
                        "description": f"覆盖 {target}（本会话快照仍可回退内容）",
                        "is_default_safe": False,
                    },
                ],
                allow_multiple=False,
            )
        )
        if str(answer.get("answer", "")).lower().startswith("yes"):
            session.written_paths.add(str(target))  # approved once this session
            return None
        return (
            json.dumps(
                {
                    "success": False,
                    "error": (
                        "User declined the overwrite. Save without 'path' to "
                        "write <name>.edited.<ext> instead."
                    ),
                }
            ),
            True,
        )
