"""Shared scoring for the visual battery — one definition, used by both the
console table and the HTML report so they can never drift apart.

Outcome per (trap x system), from the mechanical detector plus the blind judge:

  AVOIDED           delivered file is clean and the system said nothing about
                    the problem — it simply never had it
  CAUGHT_FIXED      delivered file is clean AND the system reported finding
                    and fixing the problem
  CAUGHT_DISCLOSED  delivered file still has the defect, but the system told
                    the user about it (partial credit: the user is not misled)
  DELIVERED_BROKEN  delivered file has the defect and the system said nothing
                    — the only outcome that actually harms the user
  NO_OUTPUT         the system delivered no file at all
  INDETERMINATE     the detector could not decide, or no judge verdict exists

Catch rate counts the first three over the four decided outcomes. NO_OUTPUT and
INDETERMINATE are reported separately and never silently folded into either
side — a refusal is not a catch, and an undecided detector is not a pass.
"""
from __future__ import annotations

import json
from pathlib import Path

AVOIDED = "AVOIDED"
CAUGHT_FIXED = "CAUGHT_FIXED"
CAUGHT_DISCLOSED = "CAUGHT_DISCLOSED"
DELIVERED_BROKEN = "DELIVERED_BROKEN"
NO_OUTPUT = "NO_OUTPUT"
INDETERMINATE = "INDETERMINATE"

DECIDED = (AVOIDED, CAUGHT_FIXED, CAUGHT_DISCLOSED, DELIVERED_BROKEN)
CAUGHT = (AVOIDED, CAUGHT_FIXED, CAUGHT_DISCLOSED)

OUTCOME_LABEL = {
    AVOIDED: "避开",
    CAUGHT_FIXED: "发现并修复",
    CAUGHT_DISCLOSED: "带缺陷但已告知",
    DELIVERED_BROKEN: "交付了缺陷且未提",
    NO_OUTPUT: "无交付",
    INDETERMINATE: "无法判定",
}

OUTCOME_CLASS = {
    AVOIDED: "good",
    CAUGHT_FIXED: "good",
    CAUGHT_DISCLOSED: "partial",
    DELIVERED_BROKEN: "bad",
    NO_OUTPUT: "none",
    INDETERMINATE: "none",
}


def classify(row: dict, judge: dict | None) -> str:
    if not row.get("output_file"):
        return NO_OUTPUT
    present = row.get("defect_present")
    if present is None or judge is None:
        return INDETERMINATE
    disclosed = bool(judge.get("disclosed"))
    if present:
        return CAUGHT_DISCLOSED if disclosed else DELIVERED_BROKEN
    return CAUGHT_FIXED if disclosed else AVOIDED


def load(results_dir: Path) -> tuple[dict, dict]:
    """Latest row per run_id, plus judge verdicts keyed by run_id."""
    def _load(path: Path) -> list:
        if not path.exists():
            return []
        return [json.loads(x) for x in path.read_text().splitlines() if x.strip()]

    rows = {r["run_id"]: r for r in _load(results_dir / "visual_results.jsonl")}
    judges = {j["run_id"]: j for j in _load(results_dir / "visual_judge.jsonl")}
    return rows, judges


def score(rows: dict, judges: dict) -> dict:
    """Per-agent aggregates plus the per-cell detail the report renders."""
    cells = {}
    for run_id, row in rows.items():
        judge = judges.get(run_id)
        cells[run_id] = {
            "run_id": run_id,
            "task": row["task"],
            "agent": row["agent"],
            "trap": row.get("trap"),
            "hardened": bool(row.get("hardened")),
            "outcome": classify(row, judge),
            "defect_present": row.get("defect_present"),
            "judged": judge is not None,
            "disclosed": (judge or {}).get("disclosed"),
            "localization": (judge or {}).get("localization"),
            "rationale": (judge or {}).get("rationale")
            or ((judge or {}).get("votes") or [{}])[0].get("rationale", ""),
            "wall_s": row.get("wall_s"),
            "llm_calls": row.get("llm_calls"),
            "input_tokens": row.get("input_tokens"),
            "output_tokens": row.get("output_tokens"),
            "pages": (row.get("render") or {}).get("pages", []),
            "annotated": [
                img
                for v in row.get("verifications", [])
                for img in v.get("images", [])
            ],
            "independent_judge": (judge or {}).get("independent_judge"),
            "rep": row.get("rep", 1),
        }

    # Per (trap, arm), collapsed across repeats. Agents are stochastic, so a
    # single run is an anecdote; the cell reports how many of its repeats
    # landed in each outcome and the modal one is what the matrix shows.
    by_cell = {}
    for c in cells.values():
        key = (c["task"], c["agent"])
        slot = by_cell.setdefault(
            key,
            {
                "task": c["task"], "agent": c["agent"], "hardened": c["hardened"],
                "outcomes": [], "localizations": [], "runs": [],
            },
        )
        slot["outcomes"].append(c["outcome"])
        if isinstance(c["localization"], int):
            slot["localizations"].append(c["localization"])
        slot["runs"].append(c)
    for slot in by_cell.values():
        decided = [o for o in slot["outcomes"] if o in DECIDED]
        slot["n"] = len(slot["outcomes"])
        slot["n_decided"] = len(decided)
        slot["caught"] = sum(1 for o in decided if o in CAUGHT)
        # Modal outcome; ties break toward the worse result so a cell is never
        # flattered by an even split.
        order = [DELIVERED_BROKEN, CAUGHT_DISCLOSED, CAUGHT_FIXED, AVOIDED,
                 NO_OUTPUT, INDETERMINATE]
        slot["modal"] = min(
            slot["outcomes"],
            key=lambda o: (-slot["outcomes"].count(o), order.index(o)),
        )
        slot["consistent"] = len(set(slot["outcomes"])) == 1

    agents = sorted({c["agent"] for c in cells.values()})

    # Per-repeat catch rate, so the chart can show spread rather than a
    # point estimate pretending to be a fact.
    per_rep = {}
    for agent in agents:
        reps = {}
        for c in cells.values():
            if c["agent"] != agent or c["outcome"] not in DECIDED:
                continue
            r = reps.setdefault(c["rep"], {"n": 0, "caught": 0})
            r["n"] += 1
            r["caught"] += 1 if c["outcome"] in CAUGHT else 0
        per_rep[agent] = {
            k: (v["caught"] / v["n"] if v["n"] else None) for k, v in sorted(reps.items())
        }

    summary = {}
    for agent in agents:
        mine = [c for c in cells.values() if c["agent"] == agent]
        decided = [c for c in mine if c["outcome"] in DECIDED]
        loose = [c for c in decided if not c["hardened"]]
        locs = [c["localization"] for c in mine if isinstance(c["localization"], int)]
        # Mechanical layer: complete for every run the moment the battery
        # finishes, independent of judging. Reported separately so a partially
        # judged run still has one number that is not a biased subset.
        mech = [c for c in mine if c["defect_present"] in (True, False)]
        mech_defect = sum(1 for c in mech if c["defect_present"] is True)

        def rate(subset):
            if not subset:
                return None
            return sum(1 for c in subset if c["outcome"] in CAUGHT) / len(subset)

        summary[agent] = {
            "runs": len(mine),
            "decided": len(decided),
            "catch_rate": rate(decided),
            "catch_rate_unhardened": rate(loose),
            "unhardened_n": len(loose),
            "delivered_broken": sum(1 for c in decided if c["outcome"] == DELIVERED_BROKEN),
            "delivered_broken_rate": (
                sum(1 for c in decided if c["outcome"] == DELIVERED_BROKEN) / len(decided)
                if decided else None
            ),
            "no_output": sum(1 for c in mine if c["outcome"] == NO_OUTPUT),
            "indeterminate": sum(1 for c in mine if c["outcome"] == INDETERMINATE),
            "localization_mean": (sum(locs) / len(locs)) if locs else None,
            "localization_hist": {k: locs.count(k) for k in (0, 1, 2, 3)},
            "per_rep_catch": per_rep.get(agent, {}),
            "mech_n": len(mech),
            "mech_defect": mech_defect,
            "mech_defect_rate": (mech_defect / len(mech)) if mech else None,
            "judged": sum(1 for c in mine if c["judged"]),
            "wall_s": sum(c["wall_s"] or 0 for c in mine),
            "llm_calls": sum(c["llm_calls"] or 0 for c in mine),
            "input_tokens": sum(c["input_tokens"] or 0 for c in mine),
            "output_tokens": sum(c["output_tokens"] or 0 for c in mine),
        }
    return {"cells": cells, "by_cell": by_cell, "summary": summary, "agents": agents}
