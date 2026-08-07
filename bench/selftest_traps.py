"""Trap calibration.

For every trap, three states must hold:

  clean   — the freshly built fixture does NOT trip the detector
  naive   — a plausible byte-level edit DOES trip it
  correct — the edit a system that noticed would make does NOT trip it

A trap that fires on the untouched fixture is a false-positive generator. A
trap that never fires measures nothing. A trap that fires on the correct edit
too is unfair. All three have to pass before spending a cent of API budget.

  python selftest_traps.py                 # VT1-VT5, offline
  python selftest_traps.py --with-render   # + VT6, drives real Word
"""
import argparse
import shutil
import sys
import tempfile
from pathlib import Path

BENCH = Path(__file__).resolve().parent
sys.path.insert(0, str(BENCH))

import trap_detectors as td  # noqa: E402
import visual_fixtures as vf  # noqa: E402

# fixture -> trap id. VT6 is render-only and handled separately.
STRUCTURAL_CASES = [
    ("vis_cjk.docx", "cjk_font_slot", {}),
    ("vis_narrow.xlsx", "col_overflow", {}),
    ("vis_table.docx", "table_overflow", {}),
    ("vis_chart.xlsx", "chart_overlap", {}),
    ("vis_contrast.xlsx", "contrast", {}),
    ("vis_rowclip.xlsx", "row_clip", {}),
    ("vis_merge.xlsx", "header_lost", {"labels": vf.MERGE_HEADERS}),
]

# Render-only traps: (fixture, trap, detector kwargs)
RENDER_CASES = [
    ("vis_orphan.docx", "orphan_heading", {"heading": vf.ORPHAN_HEADING}),
    ("vis_wide.xlsx", "print_split", {"labels": vf.WIDE_HEADERS}),
]


def _stage(tmp: Path, fixture: str, mutate=None) -> Path:
    path = vf.build(fixture, tmp)
    if mutate:
        mutate(path)
    return path


def check_structural(verbose: bool = True) -> list[str]:
    failures = []
    for fixture, trap, kwargs in STRUCTURAL_CASES:
        with tempfile.TemporaryDirectory() as td_:
            tmp = Path(td_)
            states = {
                "clean": (None, False),
                "naive": (vf.NAIVE_EDITS[fixture], True),
                "correct": (vf.CORRECT_EDITS[fixture], False),
            }
            for state, (mutate, expected) in states.items():
                out = _stage(tmp, fixture, mutate)
                res = td.run_detector(trap, out, **kwargs)
                got = res["present"]
                ok = got is expected
                if not ok:
                    failures.append(
                        f"{trap}[{state}]: expected present={expected}, got {got} "
                        f"({res['evidence']})"
                    )
                if verbose:
                    mark = "ok  " if ok else "FAIL"
                    detail = ""
                    if state == "naive" and got:
                        s = res["evidence"].get("samples") or []
                        detail = f"  ← {s[0]}" if s else ""
                    print(f"  {mark} {trap:16} {state:8} present={got}{detail}")
    return failures


def check_render_cases(verbose: bool = True, filler: int | None = None,
                       only: str | None = None) -> list[str]:
    """The render-only traps: no structural shortcut exists for either, so the
    calibration has to drive real Word.
    """
    from render_pdf import render_keep_pdf

    failures = []
    for fixture, trap, kwargs in RENDER_CASES:
        if only and trap != only:
            continue
        with tempfile.TemporaryDirectory() as td_:
            tmp = Path(td_)
            for state, mutate, expected in (
                ("clean", None, False),
                ("naive", vf.NAIVE_EDITS[fixture], True),
                ("correct", vf.CORRECT_EDITS[fixture], False),
            ):
                work = tmp / state
                work.mkdir()
                out = work / fixture
                if fixture == "vis_orphan.docx" and filler:
                    vf.vis_orphan(out, filler=filler)
                else:
                    vf.BUILDERS[fixture](out)
                if mutate:
                    mutate(out)
                r = render_keep_pdf(out, work / "render")
                if not r.ok or r.pdf is None:
                    failures.append(f"{trap}[{state}]: render failed ({r.error})")
                    continue
                res = td.run_detector(trap, out, pdf=r.pdf, **kwargs)
                got = res["present"]
                ok = got is expected
                if not ok:
                    failures.append(
                        f"{trap}[{state}]: expected present={expected}, got {got} "
                        f"({res['evidence']})"
                    )
                if verbose:
                    ev = {k: v for k, v in res["evidence"].items() if k != "samples"}
                    print(f"  {'ok  ' if ok else 'FAIL'} {trap:16} {state:8} "
                          f"present={got} {ev}")
    return failures


# Back-compat alias for the VT6 filler sweep.
def check_orphan(verbose: bool = True, filler: int | None = None) -> list[str]:
    return check_render_cases(verbose=verbose, filler=filler, only="orphan_heading")


def calibrate_vt6(lo: int = 5, hi: int = 16) -> None:
    """Sweep the filler length to find a value where the trap springs cleanly.

    Word decides pagination, not us — the only way to place the heading near a
    page foot is to measure. Prints the first filler count where
    clean=False / naive=True / correct=False all hold.
    """
    for n in range(lo, hi + 1):
        print(f"\nfiller={n}")
        fails = check_orphan(verbose=True, filler=n)
        if not fails:
            print(f"\n>>> filler={n} satisfies all three states. "
                  f"Set ORPHAN_FILLER_PARAGRAPHS = {n} in visual_fixtures.py")
            return
    print("\nno filler count in range produced a clean trap; widen the sweep "
          "or restructure the fixture")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--with-render", action="store_true",
                    help="also calibrate VT6 (drives real Microsoft Word)")
    ap.add_argument("--calibrate-vt6", action="store_true",
                    help="sweep VT6's filler length to find a working trap")
    args = ap.parse_args()

    if args.calibrate_vt6:
        calibrate_vt6()
        return

    print("calibrating structural traps (VT1-VT5)...")
    failures = check_structural()
    if args.with_render:
        print("calibrating render-only traps (VT6/VT9, real Word)...")
        failures += check_render_cases()
    else:
        print("skipping VT6/VT9 (need --with-render and a real Word install)")

    if failures:
        print(f"\n{len(failures)} calibration failure(s):")
        for f in failures:
            print(f"  - {f}")
        sys.exit(1)
    print("\nall traps calibrated: clean=False, naive=True, correct=False")


if __name__ == "__main__":
    main()
