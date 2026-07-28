"""Drives the real office-agent harness end-to-end and captures every artifact
the closed loop produces along the way — plan, tool calls, renders, and the
independent verifier's full structured verdicts — before the CLI's normal
session cleanup would delete them.

This bypasses cli.py on purpose: cli.py always calls
`session.ctx.sessions.close_all()` at the end, which deletes the session
working directory (render PDFs/PNGs, audit log) via `EditSession.cleanup()`.
Driving `AgentSession` directly, the way this script does, is the same public
API the CLI uses — it just skips that final teardown so we can copy the
artifacts out first.
"""
import json
import shutil
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent

from office_agent.agent import loop as loop_module
from office_agent.agent.loop import AgentSession, BaseUI
from office_agent.agent.plan import Plan
from office_agent.config import Config

INSTRUCTION = (
    "把标题《季度工作总结》加粗、字号改成24号并居中；"
    "然后在文档末尾新增一个2行3列的表格，表头为“任务、负责人、状态”，"
    "第二行填一行示例数据，并给整个表格加上黑色实线边框。"
)
INPUT_DOC = HERE / "input" / "quarterly_summary.docx"

events = []  # chronological structured event log — the machine-readable transcript
verify_results = []  # full VerificationResult payloads, not just the summary string


def _log(kind: str, **fields):
    events.append({"t": round(time.time(), 3), "kind": kind, **fields})


# ---------------------------------------------------------------- instrumentation
# Wrap verify_edit so we capture the FULL structured verdict (passed, confidence,
# problems[]) — the loop only ever forwards `.summary()` text to the UI, so this
# is the one place that sees the independent verifier's real output.
_real_verify_edit = loop_module.verify_edit


def _recording_verify_edit(
    llm, step_description, operations_summary, after_images, before_images=None,
    max_pages=4, extra_note="",
):
    result = _real_verify_edit(
        llm, step_description, operations_summary, after_images, before_images,
        max_pages, extra_note,
    )
    record = {
        "step_description": step_description,
        "operations_summary": operations_summary,
        "after_pages_shown": [img.index for img in after_images[:max_pages]],
        "before_pages_shown": [img.index for img in (before_images or [])[:max_pages]],
        "extra_note": extra_note,
        "passed": result.passed,
        "confidence": result.confidence,
        "skipped": result.skipped,
        "blocking": result.blocking,
        "problems": result.problems,
        "summary": result.summary(),
    }
    verify_results.append(record)
    _log("verify_result", **record)
    return result


loop_module.verify_edit = _recording_verify_edit


class RecordingUI(BaseUI):
    """Logs every harness<->UI event as structured JSON, then falls back to the
    same headless defaults BaseUI already provides (auto-select safe options)."""

    def emit_text(self, text: str) -> None:
        _log("assistant_text", text=text)
        print(f"\n[assistant]\n{text}\n")

    def notify(self, message: str) -> None:
        _log("notify", message=message)
        print(f"· {message}")

    def tool_call(self, name: str, tool_input) -> None:
        _log("tool_call", name=name, input=tool_input)
        print(f"→ {name}({json.dumps(tool_input, ensure_ascii=False, default=str)[:200]})")

    def tool_result(self, name: str, result) -> None:
        _log("tool_result", name=name, result=result)
        ok = result.get("success", True)
        print(f"  {'✓' if ok else '✗'} {json.dumps(result, ensure_ascii=False, default=str)[:200]}")

    def plan_update(self, plan: Plan) -> None:
        snapshot = [{"index": i, "description": s.description, "status": s.status}
                    for i, s in enumerate(plan.steps)]
        _log("plan_update", steps=snapshot)
        print("\n[plan]\n" + plan.render_text() + "\n")

    def ask_user(self, params):
        answer = super().ask_user(params)
        _log("ask_user", question=params.question,
             options=[o.label for o in (params.options or [])], answer=answer)
        print(f"? {params.question} -> {answer}")
        return answer


def main():
    config = Config(non_interactive=True, auto_approve_overwrite=True,
                     visual_verify=True, verbose=True)
    config.require_api_key()

    ui = RecordingUI()
    session = AgentSession(config, ui=ui)

    _log("turn_start", instruction=INSTRUCTION, input_doc=str(INPUT_DOC))
    result = session.run_turn(INSTRUCTION, str(INPUT_DOC))
    _log("turn_end", text=result.text, saved_paths=result.saved_paths, aborted=result.aborted)

    # ---- capture everything BEFORE any cleanup could delete it ----
    out_run = HERE / "run"
    out_renders = HERE / "renders"
    out_output = HERE / "output"
    out_run.mkdir(exist_ok=True)
    out_renders.mkdir(exist_ok=True)
    out_output.mkdir(exist_ok=True)

    with (out_run / "events.jsonl").open("w", encoding="utf-8") as f:
        for e in events:
            f.write(json.dumps(e, ensure_ascii=False, default=str) + "\n")

    with (out_run / "verify_results.json").open("w", encoding="utf-8") as f:
        json.dump(verify_results, f, ensure_ascii=False, indent=2, default=str)

    if result.plan:
        plan_snapshot = [{"index": i, "description": s.description, "status": s.status}
                          for i, s in enumerate(result.plan.steps)]
        with (out_run / "plan_final.json").open("w", encoding="utf-8") as f:
            json.dump(plan_snapshot, f, ensure_ascii=False, indent=2)

    with (out_run / "instruction.txt").open("w", encoding="utf-8") as f:
        f.write(INSTRUCTION + "\n")

    with (out_run / "turn_result.json").open("w", encoding="utf-8") as f:
        json.dump(
            {"text": result.text, "saved_paths": result.saved_paths, "aborted": result.aborted},
            f, ensure_ascii=False, indent=2,
        )

    # Copy each render directory (one per distinct rendered content hash, in
    # chronological order) — each holds render.pdf + page_N.png, plus
    # annotated_page_N.png for any page the verifier was actually shown a
    # red-boxed diff crop of.
    doc_ids = list(session.ctx.sessions.sessions.keys())
    _log("doc_ids", doc_ids=doc_ids)
    for doc_id in doc_ids:
        edit_session = session.ctx.sessions.sessions[doc_id]
        renderer = session.renderers.get(doc_id)
        if renderer is None:
            continue
        render_dirs = sorted(
            (d for d in renderer.render_dir.iterdir() if d.is_dir()),
            key=lambda d: d.stat().st_mtime,
        )
        for i, d in enumerate(render_dirs, start=1):
            dest = out_renders / f"step-{i}"
            if dest.exists():
                shutil.rmtree(dest)
            shutil.copytree(d, dest)
            print(f"copied render dir {d} -> {dest}")

        # session-level audit trail (every successful mutating tool call)
        if edit_session.audit_path.exists():
            shutil.copy2(edit_session.audit_path, HERE / "audit.jsonl")

    # The saved "after" document, if the model saved one.
    for p in result.saved_paths:
        src = Path(p)
        if src.exists():
            shutil.copy2(src, out_output / src.name)
            print(f"copied saved output {src} -> {out_output / src.name}")

    print("\n=== DONE ===")
    print("events:", len(events), " verify calls:", len(verify_results))
    print("saved_paths:", result.saved_paths)
    print("aborted:", result.aborted)

    # Deliberately skip session.ctx.sessions.close_all() — we already copied
    # everything we need out of the session dir; leaving it in place costs
    # nothing (office-agent gc reclaims anything >24h old).


if __name__ == "__main__":
    sys.exit(main())
