"""Aggregate results.jsonl (+ renders/index.jsonl + judge scores if present)
into per-agent comparison tables. Mechanical aggregation only.
"""
import json
from collections import defaultdict
from pathlib import Path

BENCH = Path(__file__).resolve().parent
RESULTS = BENCH / "results"


def load_jsonl(path):
    if not Path(path).exists():
        return []
    return [json.loads(x) for x in Path(path).read_text().splitlines() if x.strip()]


def main():
    rows = load_jsonl(RESULTS / "results.jsonl")
    renders = {r["run_id"]: r for r in load_jsonl(RESULTS / "renders" / "index.jsonl")}
    judges = {j["run_id"]: j for j in load_jsonl(RESULTS / "judge_scores.jsonl")}

    by_agent = defaultdict(list)
    latest = {}
    for r in rows:  # keep the latest entry per run_id
        latest[r["run_id"]] = r
    for r in latest.values():
        by_agent[r["agent"]].append(r)

    print(f"{'':24} {'officecli_m3':>14} {'office_agent':>14}")
    agents = ["officecli", "office_agent"]

    def stat(fn, fmt="{:>14}"):
        return "".join(fmt.format(fn(by_agent[a])) for a in agents)

    n = {a: len(by_agent[a]) for a in agents}
    print(f"{'runs':24}" + stat(lambda rs: len(rs)))
    print(f"{'passed (mechanical)':24}" + stat(lambda rs: sum(1 for r in rs if r['passed'])))
    print(f"{'pass rate':24}" + stat(lambda rs: f"{sum(1 for r in rs if r['passed'])/max(1,len(rs)):.0%}"))
    print(f"{'avg wall s':24}" + stat(lambda rs: f"{sum(r['wall_s'] for r in rs)/max(1,len(rs)):.0f}"))
    print(f"{'median wall s':24}" + stat(lambda rs: f"{sorted(r['wall_s'] for r in rs)[len(rs)//2]:.0f}"))
    print(f"{'avg llm calls':24}" + stat(lambda rs: f"{sum(r['llm_calls'] for r in rs)/max(1,len(rs)):.1f}"))
    print(f"{'total input tok':24}" + stat(lambda rs: sum(r['input_tokens'] for r in rs)))
    print(f"{'total output tok':24}" + stat(lambda rs: sum(r['output_tokens'] for r in rs)))
    print(f"{'llm errors':24}" + stat(lambda rs: sum(r.get('llm_errors', 0) for r in rs)))
    print(f"{'orig preserved':24}" + stat(lambda rs: sum(1 for r in rs if r.get('original_preserved'))))
    print(
        f"{'openxml valid':24}"
        + stat(lambda rs: sum(1 for r in rs if (r.get('openxml_validate') or {}).get('exit') == 0))
    )
    print(
        f"{'render ok':24}"
        + stat(lambda rs: sum(1 for r in rs if (renders.get(r['run_id']) or {}).get('render_ok')))
    )
    if judges:
        def avg_judge(rs, key):
            vals = [judges[r["run_id"]].get(key) for r in rs if r["run_id"] in judges]
            vals = [v for v in vals if isinstance(v, (int, float))]
            return f"{sum(vals)/len(vals):.1f}" if vals else "-"
        print(f"{'judge visual avg':24}" + stat(lambda rs: avg_judge(rs, 'visual_score')))
        print(f"{'judge behavior avg':24}" + stat(lambda rs: avg_judge(rs, 'behavior_score')))

    print("\nper-task:")
    print(f"{'task':22} {'officecli':>32} {'office_agent':>32}")
    tasks = sorted({r["task"] for r in latest.values()})
    per = {(r["task"], r["agent"]): r for r in latest.values()}
    for t in tasks:
        cells = []
        for a in agents:
            r = per.get((t, a))
            if not r:
                cells.append("-")
                continue
            mark = "PASS" if r["passed"] else "FAIL"
            jt = judges.get(r["run_id"], {})
            js = f" j={jt.get('visual_score', jt.get('behavior_score', ''))}" if jt else ""
            cells.append(f"{mark} {r['wall_s']:.0f}s {r['llm_calls']}c {(r['input_tokens']+r['output_tokens'])//1000}k{js}")
        print(f"{t:22} {cells[0]:>32} {cells[1]:>32}")

    fails = [r for r in latest.values() if not r["passed"]]
    if fails:
        print("\nfailed checks:")
        for r in sorted(fails, key=lambda x: x["run_id"]):
            bad = [c for c in r["checks"] if not c["ok"]]
            print(f"  {r['run_id']}: " + "; ".join(f"{c['name']}({c['detail'][:60]})" for c in bad))


if __name__ == "__main__":
    main()
