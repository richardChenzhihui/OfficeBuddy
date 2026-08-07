"""Re-resolve, re-render and re-detect existing visual runs — no API calls.

The agents' work is already on disk under workdir_visual/ and their reports are
already in results/logs/. When the harness (not the agent) was wrong — a bad
output resolver, a fixed detector, a new trap threshold — the fix belongs here,
not in a re-run that spends tokens again to produce identical agent behaviour.

  python regrade_visual.py                 # every run whose output is missing
  python regrade_visual.py --all           # every run
  python regrade_visual.py --agent office_agent_noverify
"""
import argparse
import json
import shutil
import sys
from pathlib import Path

BENCH = Path(__file__).resolve().parent
sys.path.insert(0, str(BENCH))

import trap_detectors as td  # noqa: E402
import visual_fixtures  # noqa: E402
from run_bench import parse_report, sha  # noqa: E402
from run_visual_bench import (  # noqa: E402
    RENDERS, RESULTS_FILE, WORKROOT, resolve_output,
)
from visual_tasks import TASK_BY_ID  # noqa: E402

visual_fixtures.register()


def report_from_log(run_id: str) -> dict:
    log = BENCH / "results" / "logs" / f"{run_id}.log"
    if not log.exists():
        return {}
    return parse_report(log.read_text())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--agent", default="")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    from render_pdf import render_keep_pdf

    rows = [
        json.loads(x) for x in RESULTS_FILE.read_text().splitlines() if x.strip()
    ]
    latest = {r["run_id"]: r for r in rows}

    targets = []
    for run_id, row in latest.items():
        if args.agent and row["agent"] != args.agent:
            continue
        if args.all or not row.get("output_file"):
            targets.append(run_id)
    print(f"{len(targets)} run(s) to regrade")
    if args.dry_run:
        for t in targets:
            print("  ", t)
        return

    changed = 0
    for run_id in targets:
        row = latest[run_id]
        task = TASK_BY_ID[row["task"]]
        workdir = WORKROOT / run_id
        fixture_path = workdir / task.fixture
        if not workdir.exists():
            print(f"skip {run_id}: workdir gone (agent output not recoverable)")
            continue

        report = report_from_log(run_id)
        # The fixture on disk is post-run; an in-place editor already changed
        # it, so compare against a freshly built copy to recover orig_hash.
        probe = workdir / "__orig_probe"
        probe.mkdir(exist_ok=True)
        orig_hash = sha(visual_fixtures.build(task.fixture, probe))
        shutil.rmtree(probe, ignore_errors=True)

        output = resolve_output(row["agent"], report, fixture_path, orig_hash)
        if output is None:
            print(f"skip {run_id}: still no output")
            continue

        render = row.get("render") or {}
        pdf = render.get("pdf")
        if not (pdf and Path(pdf).exists()):
            dest = RENDERS / run_id
            if dest.exists():
                shutil.rmtree(dest)
            r = render_keep_pdf(output, dest)
            render = {
                "ok": r.ok,
                "pdf": str(r.pdf) if r.pdf else None,
                "pages": [str(p) for p in r.pages],
                "error": r.error,
            }

        detection = td.run_detector(
            task.trap,
            output,
            pdf=Path(render["pdf"]) if render.get("pdf") else None,
            **task.detector_kwargs,
        )
        row.update(
            {
                "output_file": str(output),
                "defect_present": detection["present"],
                "detection": detection,
                "render": render,
                "verifications": report.get("verifications", row.get("verifications", [])),
                "ui_stream": report.get("ui_stream", row.get("ui_stream", [])),
                "regraded": True,
            }
        )
        changed += 1
        print(f"regraded {run_id}: defect={detection['present']} render_ok={render.get('ok')}",
              flush=True)

    if changed:
        RESULTS_FILE.write_text(
            "\n".join(json.dumps(r, ensure_ascii=False) for r in latest.values()) + "\n"
        )
    print(f"{changed} row(s) updated")


if __name__ == "__main__":
    main()
