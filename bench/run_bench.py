"""Benchmark orchestrator: runs each task on each agent sequentially (real
Word/Excel rendering and clean wall-time both need serialization), grades
outputs mechanically, and records per-run metrics from the metering proxy.

Usage:
  python run_bench.py --proxy http://127.0.0.1:8791/anthropic --tasks all --agents both
"""
import argparse
import hashlib
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
import graders  # noqa: E402
from tasks import TASKS, TASK_BY_ID  # noqa: E402

RESULTS = BENCH / "results"
WORKROOT = BENCH / "workdir"
RUN_TIMEOUT = 1800


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def proxy_stats(run_id: str, since_ts: float = 0.0) -> dict:
    stats = {
        "llm_calls": 0, "input_tokens": 0, "output_tokens": 0,
        "cache_read_tokens": 0, "llm_time_s": 0.0, "llm_errors": 0,
    }
    log = RESULTS / "llm_calls.jsonl"
    if not log.exists():
        return stats
    for line in log.read_text().splitlines():
        try:
            e = json.loads(line)
        except json.JSONDecodeError:
            continue
        if e.get("run_id") != run_id or e.get("ts", 0) < since_ts:
            continue  # same run_id from an earlier --force attempt doesn't count
        stats["llm_calls"] += 1
        stats["input_tokens"] += e.get("input_tokens") or 0
        stats["output_tokens"] += e.get("output_tokens") or 0
        stats["cache_read_tokens"] += e.get("cache_read_tokens") or 0
        stats["llm_time_s"] += e.get("latency_s") or 0
        if e.get("status") != 200:
            stats["llm_errors"] += 1
    stats["llm_time_s"] = round(stats["llm_time_s"], 1)
    return stats


def validate_output(path: Path) -> dict:
    try:
        proc = subprocess.run(
            ["officecli", "validate", str(path)], capture_output=True, text=True, timeout=120
        )
        out = (proc.stdout + proc.stderr).strip()
        return {"exit": proc.returncode, "output": out[:400]}
    except Exception as exc:
        return {"exit": -1, "output": repr(exc)[:200]}


def parse_report(stdout: str) -> dict:
    for line in reversed(stdout.splitlines()):
        if line.startswith("===BENCH_RESULT==="):
            try:
                return json.loads(line[len("===BENCH_RESULT===") :])
            except json.JSONDecodeError:
                return {}
    return {}


def run_one(task, agent: str, proxy_url: str, force: bool) -> dict:
    run_id = f"{task.id}__{agent}"
    done_ids = set()
    results_file = RESULTS / "results.jsonl"
    if results_file.exists():
        for line in results_file.read_text().splitlines():
            try:
                done_ids.add(json.loads(line)["run_id"])
            except Exception:
                pass
    if run_id in done_ids and not force:
        print(f"skip {run_id} (already done)")
        return {}
    if run_id in done_ids and force:  # drop stale lines so stats never double-count
        lines = [
            ln
            for ln in results_file.read_text().splitlines()
            if json.loads(ln).get("run_id") != run_id
        ]
        results_file.write_text("\n".join(lines) + ("\n" if lines else ""))

    workdir = WORKROOT / run_id
    if workdir.exists():
        shutil.rmtree(workdir)
    workdir.mkdir(parents=True)
    fixture_path = fixtures.build(task.fixture, workdir)
    for extra in task.extra_fixtures:
        fixtures.build(extra, workdir)
    orig_hash = sha(fixture_path)
    (RESULTS / "current_run.txt").write_text(run_id)

    if agent == "officecli":
        cmd = [
            sys.executable, str(BENCH / "officecli_agent.py"),
            "--instruction", task.instruction, "--file", task.fixture,
            "--workdir", str(workdir), "--base-url", proxy_url,
        ]
    else:
        cmd = [
            sys.executable, str(BENCH / "office_agent_driver.py"),
            "--instruction", task.instruction, "--file", str(fixture_path),
            "--base-url", proxy_url,
        ]

    t0 = time.time()
    try:
        proc = subprocess.run(
            cmd, cwd=str(workdir), capture_output=True, text=True,
            timeout=RUN_TIMEOUT, env=os.environ.copy(),
        )
        stdout, stderr, exit_code = proc.stdout, proc.stderr, proc.returncode
    except subprocess.TimeoutExpired as exc:
        stdout = (exc.stdout or b"").decode() if isinstance(exc.stdout, bytes) else (exc.stdout or "")
        stderr = "TIMEOUT"
        exit_code = -9
    wall = round(time.time() - t0, 1)

    logs = RESULTS / "logs"
    logs.mkdir(exist_ok=True)
    (logs / f"{run_id}.log").write_text(
        f"### CMD: {cmd}\n### STDOUT\n{stdout}\n### STDERR\n{stderr}\n"
    )

    report = parse_report(stdout)

    # Resolve output file
    output: Path | None = None
    if agent == "office_agent":
        # last existing save wins (saved_paths is append-ordered)
        for p in reversed(report.get("saved_paths", [])):
            if Path(p).exists():
                output = Path(p)
                break
        if output is None:
            cand = fixture_path.with_name(f"{fixture_path.stem}.edited{fixture_path.suffix}")
            if cand.exists():
                output = cand
        if output is None and sha(fixture_path) != orig_hash:
            output = fixture_path
    else:
        if sha(fixture_path) != orig_hash:
            output = fixture_path

    grade = graders.grade(task.id, output, workdir)
    validation = validate_output(output) if output else None

    result = {
        "run_id": run_id,
        "task": task.id,
        "kind": task.kind,
        "agent": agent,
        "behavioral": task.behavioral,
        "wall_s": wall,
        "exit_code": exit_code,
        "finish": report.get("finish"),
        "output_file": str(output) if output else None,
        "original_preserved": sha(fixture_path) == orig_hash if fixture_path.exists() else False,
        "passed": grade["passed"],
        "checks": grade["checks"],
        "openxml_validate": validation,
        **proxy_stats(run_id, since_ts=t0),
        "log": str(logs / f"{run_id}.log"),
    }
    with results_file.open("a") as f:
        f.write(json.dumps(result, ensure_ascii=False) + "\n")
    mark = "✅" if result["passed"] else "❌"
    print(
        f"{mark} {run_id}: {wall}s, {result['llm_calls']} llm calls, "
        f"{result['input_tokens']}+{result['output_tokens']} tok, exit={exit_code}",
        flush=True,
    )
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--proxy", default="http://127.0.0.1:8791/anthropic")
    ap.add_argument("--tasks", default="all")
    ap.add_argument("--agents", default="both")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    RESULTS.mkdir(exist_ok=True)
    WORKROOT.mkdir(exist_ok=True)
    task_list = TASKS if args.tasks == "all" else [TASK_BY_ID[t] for t in args.tasks.split(",")]
    agents = ["officecli", "office_agent"] if args.agents == "both" else [args.agents]

    print(f"battery: {len(task_list)} tasks x {agents}", flush=True)
    for task in task_list:
        for agent in agents:
            run_one(task, agent, args.proxy, args.force)
    print("battery complete", flush=True)


if __name__ == "__main__":
    main()
