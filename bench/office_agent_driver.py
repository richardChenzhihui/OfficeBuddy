"""Thin driver for office_agent so the benchmark can (a) route it through the
metering proxy via Config.base_url and (b) run it fully non-interactively —
without modifying office_agent's source.

It also captures the two things the visual battery scores that a plain
run_turn() return value does not expose:

  * every visual-verification verdict (page / element / description / severity)
  * the annotated PNGs the verifier was shown

Both are captured by wrapping, not by patching product code: the session's own
methods are decorated in place for the duration of the run. Session directories
live inside the Office sandbox container and are destroyed on close, so the
evidence is copied out before teardown.
"""
import argparse
import json
import shutil
import time
from pathlib import Path


class RecordingUI:
    """Delegates to the real CLI UI while recording everything the user would
    have seen. That stream is the input to the defect-localization judge — the
    same thing a human operator would have had to work from.
    """

    def __init__(self, inner):
        self._inner = inner
        self.stream = []

    def emit_text(self, text):
        self.stream.append({"kind": "text", "text": text})
        return self._inner.emit_text(text)

    def notify(self, message):
        self.stream.append({"kind": "notify", "text": message})
        return self._inner.notify(message)

    def __getattr__(self, name):
        return getattr(self._inner, name)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--instruction", required=True)
    ap.add_argument("--file", required=True)
    ap.add_argument("--base-url", default="https://api.minimaxi.com/anthropic")
    ap.add_argument("--no-visual-verify", action="store_true")
    ap.add_argument("--evidence-dir", default="")
    args = ap.parse_args()

    from office_agent.agent.loop import AgentSession
    from office_agent.config import Config
    from office_agent.ui.cli_ui import CliUI

    config = Config(
        base_url=args.base_url,
        visual_verify=not args.no_visual_verify,
        non_interactive=True,
        verbose=True,
    )
    config.require_api_key()
    ui = RecordingUI(CliUI(verbose=True, non_interactive=True))
    session = AgentSession(config, ui=ui)

    evidence_dir = Path(args.evidence_dir).resolve() if args.evidence_dir else None
    if evidence_dir:
        evidence_dir.mkdir(parents=True, exist_ok=True)

    verifications = []
    original_verify = session._verify_doc

    def recording_verify(doc_id, step_description):
        verdict = original_verify(doc_id, step_description)
        entry = {
            "doc_id": doc_id,
            "step": step_description,
            "passed": bool(verdict.passed),
            "skipped": bool(verdict.skipped),
            "blocking": bool(verdict.blocking),
            "confidence": round(float(verdict.confidence or 0.0), 3),
            "problems": list(verdict.problems or []),
            "images": [],
        }
        # Copy the annotated pages out now: the session directory is inside the
        # Office sandbox container and is deleted at close_all().
        if evidence_dir:
            try:
                render_dir = session.ctx.sessions.get(doc_id).session_dir / "render"
                for png in sorted(render_dir.glob("**/annotated_*.png")):
                    target = evidence_dir / f"v{len(verifications):02d}_{png.name}"
                    if not target.exists():
                        shutil.copy(png, target)
                        entry["images"].append(str(target))
            except Exception as exc:
                entry["image_error"] = repr(exc)[:200]
        verifications.append(entry)
        return verdict

    session._verify_doc = recording_verify

    t0 = time.time()
    status = "done"
    text = ""
    saved = []
    try:
        result = session.run_turn(args.instruction, args.file)
        text = (result.text or "")[:4000]
        saved = [p for p in (result.saved_paths or []) if p]
        if getattr(result, "aborted", False):
            status = "aborted"
    except Exception as exc:
        status = "error"
        text = repr(exc)[:4000]
    finally:
        try:
            session.ctx.sessions.close_all()
        except Exception:
            pass

    report = {
        "agent": "office_agent",
        "wall_s": round(time.time() - t0, 1),
        "finish": {"status": status, "summary": text},
        "saved_paths": saved,
        "verifications": verifications,
        "ui_stream": ui.stream[-80:],
        "evidence_dir": str(evidence_dir) if evidence_dir else None,
    }
    print("\n===BENCH_RESULT===" + json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
