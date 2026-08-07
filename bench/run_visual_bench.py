"""Orchestrator for the visual-trap battery (M1 catch rate, M3 localization).

Reuses run_bench's metering/spawn helpers rather than duplicating them; the
original 20-task battery is left untouched so its results stay reproducible.

Per run:
  1. build the fixture fresh (byte-identical for every agent)
  2. spawn the agent as a subprocess through the metering proxy
  3. resolve the delivered file
  4. render it through the REAL Word/Excel and keep the PDF + page PNGs
  5. run the trap's mechanical detector on the delivered bytes (and, for the
     render-only trap, on the PDF)

Whether the system *told the user* about a defect is deliberately not decided
here — that is a semantic judgement and belongs to judge_visual.py.

  python run_visual_bench.py --agents office_agent,officecli_noshot
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

BENCH = Path(__file__).resolve().parent
sys.path.insert(0, str(BENCH))

import fixtures  # noqa: E402
import trap_detectors as td  # noqa: E402
import visual_fixtures  # noqa: E402
from run_bench import parse_report, proxy_stats, sha  # noqa: E402
from visual_tasks import TASK_BY_ID, VISUAL_TASKS  # noqa: E402

visual_fixtures.register()

RESULTS = BENCH / "results"
WORKROOT = BENCH / "workdir_visual"
RENDERS = RESULTS / "visual_renders"
RESULTS_FILE = RESULTS / "visual_results.jsonl"
RUN_TIMEOUT = 1800

# agent label -> (script, extra args). The officecli screenshot ablation is a
# separate labelled arm, never a silent change to the default one.
# The 2x2: {our harness, their CLI} x {can look, cannot look}. Running all
# four turns "we beat them" into "here is what looking buys, for both systems"
# — which is the claim the render loop actually supports.
AGENTS = {
    "office_agent": ("office_agent_driver.py", []),
    "office_agent_noverify": ("office_agent_driver.py", ["--no-visual-verify"]),
    "officecli": ("officecli_agent.py", []),
    "officecli_noshot": ("officecli_agent.py", ["--no-screenshot"]),
}


def build_cmd(agent, task, workdir, fixture_path, proxy_url):
    script, extra = AGENTS[agent]
    if agent.startswith("officecli"):
        return [
            sys.executable, str(BENCH / script),
            "--instruction", task.instruction,
            "--file", task.fixture,
            "--workdir", str(workdir),
            "--base-url", proxy_url,
            *extra,
        ]
    return [
        sys.executable, str(BENCH / script),
        "--instruction", task.instruction,
        "--file", str(fixture_path),
        "--base-url", proxy_url,
        "--evidence-dir", str(workdir / "evidence"),
        *extra,
    ]


def resolve_output(agent, report, fixture_path, orig_hash):
    """Which file did this agent actually deliver?

    Match the family, not the exact arm name: office_agent_noverify saves
    copy-on-write exactly like office_agent does, and an `== "office_agent"`
    test silently threw away every one of its outputs.
    """
    if agent.startswith("office_agent"):
        for p in reversed(report.get("saved_paths", [])):
            if Path(p).exists():
                return Path(p)
        cand = fixture_path.with_name(f"{fixture_path.stem}.edited{fixture_path.suffix}")
        if cand.exists():
            return cand
    if fixture_path.exists() and sha(fixture_path) != orig_hash:
        return fixture_path  # edited in place
    return None


def load_done() -> set:
    if not RESULTS_FILE.exists():
        return set()
    out = set()
    for line in RESULTS_FILE.read_text().splitlines():
        try:
            out.add(json.loads(line)["run_id"])
        except Exception:
            pass
    return out


def drop_run(run_id: str) -> None:
    if not RESULTS_FILE.exists():
        return
    lines = [
        ln
        for ln in RESULTS_FILE.read_text().splitlines()
        if ln.strip() and json.loads(ln).get("run_id") != run_id
    ]
    RESULTS_FILE.write_text("\n".join(lines) + ("\n" if lines else ""))


def run_one(task, agent: str, proxy_url: str, force: bool, rep: int = 1) -> dict:
    from render_pdf import render_keep_pdf

    # rep 1 keeps the historical bare run_id so earlier results stay addressable.
    run_id = f"{task.id}__{agent}" + (f"__r{rep}" if rep > 1 else "")
    done = load_done()
    if run_id in done and not force:
        print(f"skip {run_id} (already done)")
        return {}
    if run_id in done:
        drop_run(run_id)

    workdir = WORKROOT / run_id
    if workdir.exists():
        shutil.rmtree(workdir)
    workdir.mkdir(parents=True)
    fixture_path = fixtures.build(task.fixture, workdir)
    orig_hash = sha(fixture_path)
    (RESULTS / "current_run.txt").write_text(run_id)

    cmd = build_cmd(agent, task, workdir, fixture_path, proxy_url)
    t0 = time.time()
    try:
        proc = subprocess.run(
            cmd, cwd=str(workdir), capture_output=True, text=True,
            timeout=RUN_TIMEOUT, env=os.environ.copy(),
        )
        stdout, stderr, exit_code = proc.stdout, proc.stderr, proc.returncode
    except subprocess.TimeoutExpired as exc:
        raw = exc.stdout or b""
        stdout = raw.decode() if isinstance(raw, bytes) else raw
        stderr, exit_code = "TIMEOUT", -9
    wall = round(time.time() - t0, 1)

    logs = RESULTS / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    (logs / f"{run_id}.log").write_text(
        f"### CMD: {cmd}\n### STDOUT\n{stdout}\n### STDERR\n{stderr}\n"
    )

    report = parse_report(stdout)
    output = resolve_output(agent, report, fixture_path, orig_hash)

    # Render the delivered bytes through the real Office app.
    render = {"ok": False, "pdf": None, "pages": [], "error": "no output file"}
    if output is not None:
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
        pdf=Path(render["pdf"]) if render["pdf"] else None,
        **task.detector_kwargs,
    )

    result = {
        "run_id": run_id,
        "task": task.id,
        "kind": task.kind,
        "agent": agent,
        "trap": task.trap,
        "hardened": task.hardened,
        "rep": rep,
        "wall_s": wall,
        "exit_code": exit_code,
        "finish": report.get("finish"),
        "output_file": str(output) if output else None,
        "original_preserved": (
            sha(fixture_path) == orig_hash if fixture_path.exists() else False
        ),
        "defect_present": detection["present"],
        "detection": detection,
        "render": render,
        # M3 inputs: what the user would actually have seen.
        "verifications": report.get("verifications", []),
        "ui_stream": report.get("ui_stream", []),
        "evidence_dir": report.get("evidence_dir"),
        "screenshots_available": report.get("screenshots_available"),
        **proxy_stats(run_id, since_ts=t0),
        "log": str(logs / f"{run_id}.log"),
    }
    RESULTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with RESULTS_FILE.open("a") as f:
        f.write(json.dumps(result, ensure_ascii=False) + "\n")

    present = detection["present"]
    mark = {True: "DEFECT", False: "clean ", None: "  ?   "}[present]
    print(
        f"[{mark}] {run_id}: {wall}s, {result['llm_calls']} llm calls, "
        f"{result['input_tokens']}+{result['output_tokens']} tok, "
        f"render_ok={render['ok']}",
        flush=True,
    )
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--proxy", default="http://127.0.0.1:8791/anthropic")
    ap.add_argument("--tasks", default="all")
    ap.add_argument("--agents",
                    default="office_agent,office_agent_noverify,officecli,officecli_noshot")
    ap.add_argument("--reps", type=int, default=1,
                    help="repeats per cell; agents are stochastic, n=1 has no error bars")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    RESULTS.mkdir(parents=True, exist_ok=True)
    WORKROOT.mkdir(parents=True, exist_ok=True)
    RENDERS.mkdir(parents=True, exist_ok=True)

    tasks = (
        VISUAL_TASKS
        if args.tasks == "all"
        else [TASK_BY_ID[t] for t in args.tasks.split(",")]
    )
    agents = [a.strip() for a in args.agents.split(",") if a.strip()]
    unknown = [a for a in agents if a not in AGENTS]
    if unknown:
        ap.error(f"unknown agent(s): {unknown}; choose from {sorted(AGENTS)}")

    total = len(tasks) * len(agents) * args.reps
    print(f"visual battery: {len(tasks)} traps x {len(agents)} arms x {args.reps} reps "
          f"= {total} runs", flush=True)
    if "officecli_noshot" in agents:
        print(
            "NOTE: officecli_noshot ablates officecli's screenshot renderer. "
            "Any number from this arm must be reported as such.",
            flush=True,
        )
    # Rep-major so a partial run still covers every cell once.
    for rep in range(1, args.reps + 1):
        for task in tasks:
            for agent in agents:
                run_one(task, agent, args.proxy, args.force, rep=rep)
    print("visual battery complete", flush=True)


if __name__ == "__main__":
    main()
