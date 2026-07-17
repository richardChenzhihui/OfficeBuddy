"""Offline tests of the agent loop control flow, driven by FakeLLM."""
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from fake_llm import FakeLLM, FakeMessage, text_block, tool_use, verify_verdict

from office_agent.agent.loop import AgentSession, BaseUI
from office_agent.agent.verifier import VerificationResult
from office_agent.config import Config
from office_agent.tools.interaction_tools import AskUserInput


class RecordingUI(BaseUI):
    def __init__(self, answers=None):
        self.texts = []
        self.notices = []
        self.questions = []
        self.plans = []
        self.answers = list(answers or [])

    def emit_text(self, text):
        self.texts.append(text)

    def notify(self, message):
        self.notices.append(message)

    def plan_update(self, plan):
        self.plans.append(plan.render_text())

    def ask_user(self, params: AskUserInput):
        self.questions.append(params)
        if self.answers:
            return {"answer": self.answers.pop(0)}
        return super().ask_user(params)


def make_session(script, verify_script=None, ui=None, **config_kwargs):
    config = Config(api_key="fake", **config_kwargs)
    llm = FakeLLM(script=script, verify_script=verify_script or [])
    session = AgentSession(config, ui=ui or RecordingUI(), llm=llm)
    return session, llm


def last_tool_results(llm):
    """Extract the most recent tool_result contents sent back to the model."""
    for msg in reversed(llm.calls[-1]["messages"]):
        if msg["role"] == "user":
            content = msg["content"]
            if isinstance(content, list) and content and content[0].get("type") == "tool_result":
                return content
    return []


def test_simple_text_turn():
    session, _ = make_session([FakeMessage(content=[text_block("Hi there")])])
    result = session.run_turn("hello")
    assert result.text == "Hi there"


def test_plan_and_ask_user_flow(word_doc_path):
    ui = RecordingUI(answers=["第一段"])
    session, llm = make_session(
        [
            FakeMessage(
                content=[
                    tool_use(
                        "ask_user",
                        {
                            "question": "改哪一段？",
                            "kind": "multiple_choice",
                            "options": [
                                {"label": "第一段", "description": "", "is_default_safe": True},
                                {"label": "全部", "description": "", "is_default_safe": False},
                            ],
                        },
                    )
                ],
                stop_reason="tool_use",
            ),
            FakeMessage(
                content=[
                    tool_use("propose_plan", {"steps": ["step A", "step B"]})
                ],
                stop_reason="tool_use",
            ),
            FakeMessage(content=[text_block("planned")]),
        ],
        ui=ui,
    )
    result = session.run_turn("把标题改一下")
    assert ui.questions, "ask_user was not routed to the UI"
    assert result.plan is not None
    assert len(result.plan.steps) == 2
    # The user's answer was fed back to the model as a tool result.
    answer_payload = json.dumps("第一段", ensure_ascii=False).strip('"')
    flat = json.dumps(llm.calls[-2]["messages"], ensure_ascii=False)
    assert answer_payload in flat


def test_execute_tool_and_finish(word_doc_path):
    session, llm = make_session(
        [
            FakeMessage(
                content=[tool_use("open_document", {"file_path": str(word_doc_path)})],
                stop_reason="tool_use",
            ),
            FakeMessage(content=[text_block("opened")]),
        ],
        visual_verify=False,
    )
    result = session.run_turn("open it")
    assert result.text == "opened"
    results = last_tool_results(llm)
    payload = json.loads(results[0]["content"])
    assert payload["success"] is True
    assert payload["doc_type"] == "word"


def test_update_plan_done_triggers_verification_and_repair(word_doc_path, monkeypatch):
    """Failing verification must bounce the step back with precise problems."""
    session, llm = make_session(
        [
            FakeMessage(
                content=[tool_use("open_document", {"file_path": str(word_doc_path)})],
                stop_reason="tool_use",
            ),
            FakeMessage(
                content=[tool_use("propose_plan", {"steps": ["edit heading"]})],
                stop_reason="tool_use",
            ),
            FakeMessage(
                content=[
                    tool_use(
                        "word_edit_text",
                        {
                            "doc_id": "DOC",
                            "selector": {"type": "paragraph", "index": 0},
                            "operation": "replace",
                            "text": "New Heading",
                        },
                    )
                ],
                stop_reason="tool_use",
            ),
            FakeMessage(
                content=[tool_use("update_plan", {"step_index": 0, "status": "done"})],
                stop_reason="tool_use",
            ),
            FakeMessage(
                content=[tool_use("update_plan", {"step_index": 0, "status": "done"})],
                stop_reason="tool_use",
            ),
            FakeMessage(content=[text_block("all fixed")]),
        ]
    )

    verdicts = [
        VerificationResult(
            passed=False,
            problems=[
                {
                    "page": 2,
                    "element_hint": "heading paragraph",
                    "description": "font is Calibri, expected Times New Roman",
                    "severity": "blocking",
                }
            ],
            confidence=0.9,
        ),
        VerificationResult(passed=True, confidence=0.95),
    ]
    def fake_verify(self, doc_id, desc):
        verdict = verdicts.pop(0)
        if verdict.passed:  # mirror real implementation: clear dirty state
            self.visual_dirty[doc_id] = set()
            self.pending_ops[doc_id] = []
        return verdict

    monkeypatch.setattr(AgentSession, "_verify_doc", fake_verify)

    # Rewrite doc_id placeholder once the real one exists.
    real_ids = {}
    original_handle = AgentSession._handle_tool

    def patched_handle(self, name, tool_input):
        if tool_input.get("doc_id") == "DOC" and real_ids:
            tool_input["doc_id"] = next(iter(real_ids))
        content, is_error = original_handle(self, name, tool_input)
        if name == "open_document" and not is_error:
            real_ids[json.loads(content)["doc_id"]] = True
        return content, is_error

    monkeypatch.setattr(AgentSession, "_handle_tool", patched_handle)

    result = session.run_turn("edit the heading")
    assert result.text == "all fixed"
    assert result.plan.steps[0].status == "done"
    # The failing verdict was surfaced to the model with page + problem detail.
    flat = json.dumps(llm.calls[-2]["messages"], ensure_ascii=False)
    assert "page 2" in flat
    assert "Times New Roman" in flat


def test_same_error_escalation_ladder(word_doc_path):
    """Two identical failures -> strategy switch; next failure -> ask_user."""
    ui = RecordingUI(answers=["Skip this step"])
    bad_edit = {
        "doc_id": "DOC",
        "selector": {"type": "text_match", "contains": "NONEXISTENT"},
        "operation": "replace",
        "text": "x",
    }
    session, llm = make_session(
        [
            FakeMessage(
                content=[tool_use("open_document", {"file_path": str(word_doc_path)})],
                stop_reason="tool_use",
            ),
            FakeMessage(content=[tool_use("word_edit_text", dict(bad_edit))], stop_reason="tool_use"),
            FakeMessage(content=[tool_use("word_edit_text", dict(bad_edit))], stop_reason="tool_use"),
            FakeMessage(content=[tool_use("word_edit_text", dict(bad_edit))], stop_reason="tool_use"),
            FakeMessage(content=[text_block("giving up per user decision")]),
        ],
        ui=ui,
        visual_verify=False,
    )

    real_id = {}
    original_handle = AgentSession._handle_tool

    def patched_handle(self, name, tool_input):
        if tool_input.get("doc_id") == "DOC" and real_id:
            tool_input["doc_id"] = next(iter(real_id))
        content, is_error = original_handle(self, name, tool_input)
        if name == "open_document" and not is_error:
            real_id[json.loads(content)["doc_id"]] = True
        return content, is_error

    import unittest.mock

    with unittest.mock.patch.object(AgentSession, "_handle_tool", patched_handle):
        session.run_turn("edit something impossible")

    # Failure 2 (same signature) must carry the strategy-switch instruction.
    all_payloads = json.dumps(llm.calls, ensure_ascii=False, default=str)
    assert "Do NOT repeat the same approach" in all_payloads
    # Failure 3 must have asked the user.
    assert ui.questions, "escalation did not reach ask_user"
    assert "Skip this step" in all_payloads


def test_overwrite_gate_non_interactive_refuses(word_doc_path):
    session, llm = make_session(
        [
            FakeMessage(
                content=[tool_use("open_document", {"file_path": str(word_doc_path)})],
                stop_reason="tool_use",
            ),
            FakeMessage(
                content=[
                    tool_use(
                        "save_document",
                        {"doc_id": "DOC", "path": str(word_doc_path), "overwrite": True},
                    )
                ],
                stop_reason="tool_use",
            ),
            FakeMessage(content=[text_block("saved elsewhere")]),
        ],
        visual_verify=False,
        non_interactive=True,
    )

    real_id = {}
    original_handle = AgentSession._handle_tool

    def patched_handle(self, name, tool_input):
        if tool_input.get("doc_id") == "DOC" and real_id:
            tool_input["doc_id"] = next(iter(real_id))
        content, is_error = original_handle(self, name, tool_input)
        if name == "open_document" and not is_error:
            real_id[json.loads(content)["doc_id"]] = True
        return content, is_error

    import unittest.mock

    with unittest.mock.patch.object(AgentSession, "_handle_tool", patched_handle):
        session.run_turn("save over the original")

    flat = json.dumps(llm.calls, ensure_ascii=False, default=str)
    assert "requires --yes" in flat
    assert word_doc_path.exists()


def test_task_budget_exhaustion_stops_editing(word_doc_path):
    from office_agent.config import Budgets

    session, llm = make_session(
        [
            # budgets must be set at construction: BudgetTracker copies limits
            FakeMessage(
                content=[tool_use("open_document", {"file_path": str(word_doc_path)})],
                stop_reason="tool_use",
            ),
            FakeMessage(
                content=[tool_use("get_structure", {"doc_id": "DOC"})],
                stop_reason="tool_use",
            ),
            FakeMessage(
                content=[tool_use("get_structure", {"doc_id": "DOC"})],
                stop_reason="tool_use",
            ),
            FakeMessage(content=[text_block("stopping")]),
        ],
        visual_verify=False,
        budgets=Budgets(max_tool_calls_per_task=2),
    )

    real_id = {}
    original_handle = AgentSession._handle_tool

    def patched_handle(self, name, tool_input):
        if tool_input.get("doc_id") == "DOC" and real_id:
            tool_input["doc_id"] = next(iter(real_id))
        content, is_error = original_handle(self, name, tool_input)
        if name == "open_document" and not is_error:
            real_id[json.loads(content)["doc_id"]] = True
        return content, is_error

    import unittest.mock

    with unittest.mock.patch.object(AgentSession, "_handle_tool", patched_handle):
        session.run_turn("do endless things")

    flat = json.dumps(llm.calls, ensure_ascii=False, default=str)
    assert "budget" in flat
    assert "exhausted" in flat
