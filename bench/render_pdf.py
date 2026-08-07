"""Render a document through real Word/Excel and KEEP the PDF.

render_outputs.render_office() deletes its staging area, so only the page PNGs
survive. The orphan-heading detector needs the PDF's text layer (block
positions per page), and the judge wants the PNGs — so this helper keeps both.

As in render_outputs, the agent's exact output bytes are staged and rendered;
no EditSession or openpyxl round trip is allowed to touch the artifact before
the fidelity check sees it.
"""
from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Render:
    ok: bool
    pdf: Path | None = None
    pages: list[Path] = field(default_factory=list)
    error: str | None = None


def render_keep_pdf(output: Path, dest: Path, timeout: float = 180.0) -> Render:
    """Export `output` via the real Office app into `dest/render.pdf` + page PNGs."""
    from office_agent.core.session import EXCEL_CONTAINER, WORD_CONTAINER
    from office_agent.render.applescript import export_docx_to_pdf, export_xlsx_to_pdf
    from office_agent.render.pdf_to_images import pdf_to_images

    output = Path(output)
    dest = Path(dest)
    is_word = output.suffix.lower() == ".docx"
    container = (WORD_CONTAINER if is_word else EXCEL_CONTAINER) / "office_agent_bench"
    workspace = container / f"{dest.parent.name}__{dest.name}"
    if workspace.exists():
        shutil.rmtree(workspace, ignore_errors=True)
    workspace.mkdir(parents=True, exist_ok=True)
    staged = workspace / output.name
    shutil.copy(output, staged)
    staged_pdf = workspace / "render.pdf"
    try:
        if is_word:
            export_docx_to_pdf(staged, staged_pdf, timeout=timeout)
        else:
            export_xlsx_to_pdf(staged, staged_pdf, timeout=timeout)
        dest.mkdir(parents=True, exist_ok=True)
        kept_pdf = dest / "render.pdf"
        shutil.copy(staged_pdf, kept_pdf)
        pages = []
        for img in pdf_to_images(staged_pdf, dest):
            target = dest / f"page_{img.index + 1}.png"
            if img.path != target:
                shutil.copy(img.path, target)
            pages.append(target)
        return Render(ok=True, pdf=kept_pdf, pages=pages)
    except Exception as exc:
        return Render(ok=False, error=repr(exc)[:300])
    finally:
        shutil.rmtree(workspace, ignore_errors=True)
