"""Post-battery fidelity render: open every produced output in the REAL
Word/Excel (via office_agent's zero-dialog render pipeline) and export page
PNGs. A file that real Office fails to render is a fidelity failure regardless
of which agent made it. pptx outputs use officecli's screenshot renderer
(office_agent has no pptx support to render with).

Writes results/renders/<run_id>/page_*.png and results/renders/index.jsonl.
"""
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

BENCH = Path(__file__).resolve().parent
RESULTS = BENCH / "results"
RENDERS = RESULTS / "renders"


def render_office(output: Path, dest: Path) -> dict:
    """Export the agent's EXACT output bytes — no EditSession/openpyxl round
    trip. (EditSession.flush() re-saves the file and the re-saved bytes can be
    rejected by real Excel — a finding of this benchmark; the fidelity pass
    must therefore never let SUT code touch the artifact.)"""
    from office_agent.core.session import EXCEL_CONTAINER, WORD_CONTAINER
    from office_agent.render.applescript import export_docx_to_pdf, export_xlsx_to_pdf
    from office_agent.render.pdf_to_images import pdf_to_images

    is_word = output.suffix.lower() == ".docx"
    container = (WORD_CONTAINER if is_word else EXCEL_CONTAINER) / "office_agent_bench"
    workspace = container / dest.name
    workspace.mkdir(parents=True, exist_ok=True)
    staged = workspace / output.name
    shutil.copy(output, staged)
    pdf = workspace / "render.pdf"
    try:
        if is_word:
            export_docx_to_pdf(staged, pdf, timeout=180.0)
        else:
            export_xlsx_to_pdf(staged, pdf, timeout=180.0)
        dest.mkdir(parents=True, exist_ok=True)
        images = pdf_to_images(pdf, dest)
        pages = []
        for img in images:
            target = dest / f"page_{img.index + 1}.png"
            if img.path != target:
                shutil.copy(img.path, target)
            pages.append(str(target))
        return {"render_ok": True, "pages": pages}
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def render_pptx(output: Path, dest: Path) -> dict:
    dest.mkdir(parents=True, exist_ok=True)
    target = dest / "slides_grid.png"
    proc = subprocess.run(
        ["officecli", "view", str(output), "screenshot", "-o", str(target), "--grid", "2"],
        capture_output=True,
        text=True,
        timeout=180,
    )
    if proc.returncode != 0 or not target.exists():
        raise RuntimeError((proc.stdout + proc.stderr)[:300])
    return {"render_ok": True, "pages": [str(target)]}


def main():
    only = set(sys.argv[1].split(",")) if len(sys.argv) > 1 else None
    RENDERS.mkdir(exist_ok=True)
    index_path = RENDERS / "index.jsonl"
    done = set()
    if index_path.exists():
        for line in index_path.read_text().splitlines():
            done.add(json.loads(line)["run_id"])

    for line in (RESULTS / "results.jsonl").read_text().splitlines():
        r = json.loads(line)
        run_id = r["run_id"]
        if run_id in done or (only and run_id not in only and r["task"] not in only):
            continue
        out = r.get("output_file")
        entry = {"run_id": run_id, "ts": time.time()}
        if not out or not Path(out).exists():
            entry.update({"render_ok": None, "pages": [], "note": "no output file"})
        else:
            t0 = time.time()
            try:
                if out.endswith(".pptx"):
                    entry.update(render_pptx(Path(out), RENDERS / run_id))
                else:
                    entry.update(render_office(Path(out), RENDERS / run_id))
            except Exception as exc:
                entry.update({"render_ok": False, "pages": [], "error": repr(exc)[:300]})
            entry["render_s"] = round(time.time() - t0, 1)
        with index_path.open("a") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        print(f"{run_id}: render_ok={entry.get('render_ok')} pages={len(entry.get('pages', []))}", flush=True)


if __name__ == "__main__":
    main()
