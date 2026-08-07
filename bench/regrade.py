"""Re-run mechanical graders over existing run outputs (no API calls).
Rewrites passed/checks in results.jsonl in place. Use after fixing a grader.
"""
import json
import sys
from pathlib import Path

BENCH = Path(__file__).resolve().parent
sys.path.insert(0, str(BENCH))
import graders  # noqa: E402

results_file = BENCH / "results" / "results.jsonl"
only = set(sys.argv[1].split(",")) if len(sys.argv) > 1 else None

lines = []
for line in results_file.read_text().splitlines():
    r = json.loads(line)
    if only is None or r["task"] in only or r["run_id"] in only:
        out = r.get("output_file")
        workdir = BENCH / "workdir" / r["run_id"]
        grade = graders.grade(r["task"], out if out and Path(out).exists() else None, workdir)
        changed = grade["passed"] != r["passed"]
        r["passed"], r["checks"] = grade["passed"], grade["checks"]
        print(f"{'CHANGED' if changed else 'same   '} {r['run_id']}: passed={r['passed']}")
    lines.append(json.dumps(r, ensure_ascii=False))
results_file.write_text("\n".join(lines) + "\n")
