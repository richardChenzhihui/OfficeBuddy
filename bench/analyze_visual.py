"""Console aggregation for the visual battery.

  python analyze_visual.py [--results-dir results]
"""
import argparse
import sys
from pathlib import Path

BENCH = Path(__file__).resolve().parent
sys.path.insert(0, str(BENCH))

from visual_scoring import OUTCOME_LABEL, load, score  # noqa: E402
from visual_tasks import VISUAL_TASKS  # noqa: E402


def pct(x):
    return "-" if x is None else f"{x:.0%}"


def num(x, fmt="{:.2f}"):
    return "-" if x is None else fmt.format(x)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--results-dir", default=str(BENCH / "results"))
    args = ap.parse_args()

    rows, judges = load(Path(args.results_dir))
    if not rows:
        sys.exit(f"no visual results in {args.results_dir}")
    data = score(rows, judges)
    agents, summary, cells = data["agents"], data["summary"], data["cells"]

    w = 22
    print(f"{'':{w}}" + "".join(f"{a:>20}" for a in agents))

    def line(label, fn):
        print(f"{label:{w}}" + "".join(f"{fn(summary[a]):>20}" for a in agents))

    line("runs", lambda s: s["runs"])
    line("decided", lambda s: s["decided"])
    line("M1 catch rate", lambda s: pct(s["catch_rate"]))
    line("  excl. hardened", lambda s: f'{pct(s["catch_rate_unhardened"])} (n={s["unhardened_n"]})')
    line("delivered broken", lambda s: f'{s["delivered_broken"]} ({pct(s["delivered_broken_rate"])})')
    line("no output", lambda s: s["no_output"])
    line("indeterminate", lambda s: s["indeterminate"])
    line("M3 localization", lambda s: num(s["localization_mean"]))
    line("  hist 0/1/2/3", lambda s: "/".join(str(s["localization_hist"][k]) for k in (0, 1, 2, 3)))
    line("total wall s", lambda s: f'{s["wall_s"]:.0f}')
    line("total llm calls", lambda s: s["llm_calls"])
    line("total input tok", lambda s: s["input_tokens"])

    print("\nper trap:")
    print(f"{'trap':22}" + "".join(f"{a:>26}" for a in agents))
    for task in VISUAL_TASKS:
        tag = f"{task.id}{'*' if task.hardened else ''}"
        parts = []
        for a in agents:
            c = cells.get(f"{task.id}__{a}")
            if not c:
                parts.append("-")
                continue
            loc = c["localization"]
            parts.append(
                f'{OUTCOME_LABEL[c["outcome"]]}'
                + (f" L{loc}" if isinstance(loc, int) else "")
            )
        print(f"{tag:22}" + "".join(f"{p:>26}" for p in parts))
    print("\n* hardened trap: this project defends against it by construction, so it "
          "measures a default rather than the render loop noticing something new.")

    non_independent = {
        j.get("provider") for j in judges.values() if not j.get("independent_judge")
    }
    if non_independent:
        print(f"\nWARNING: judged by a non-independent provider {sorted(non_independent)} "
              "— same model family as the system under test.")


if __name__ == "__main__":
    main()
