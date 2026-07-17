"""Document lifecycle tools: open, inspect, save, undo."""
from typing import Optional

from pydantic import BaseModel, Field

from .registry import REGISTRY, ToolContext


class OpenDocumentInput(BaseModel):
    file_path: str = Field(..., description="Path to a .docx / .xlsx / .xlsm file")


class GetStructureInput(BaseModel):
    doc_id: str = Field(..., description="Document id returned by open_document")
    max_paragraph_chars: int = Field(
        0, description="Truncate paragraph text to this many chars (0 = full text)"
    )
    include_table_cells: bool = Field(
        True, description="Include table cell contents (Word)"
    )


class SaveDocumentInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    path: Optional[str] = Field(
        None,
        description=(
            "Destination path. Omit to save next to the original as "
            "<name>.edited.<ext>. Saving over the original requires overwrite=true."
        ),
    )
    overwrite: bool = Field(
        False,
        description=(
            "Must be true to overwrite ANY pre-existing file (the original or "
            "another existing path). Requires user approval."
        ),
    )


class UndoInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    steps: int = Field(1, description="How many edit steps to undo")


class ListSnapshotsInput(BaseModel):
    doc_id: str = Field(..., description="Document id")


class RestoreSnapshotInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    snapshot_id: str = Field(..., description="Snapshot id from list_snapshots")


def _word_structure(session, max_chars: int, include_cells: bool) -> dict:
    doc = session.doc
    paragraphs = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text if max_chars <= 0 else para.text[:max_chars]
        paragraphs.append(
            {
                "index": i,
                "text": text,
                "style": para.style.name if para.style else None,
                "runs": len(para.runs),
            }
        )
    tables = []
    for i, table in enumerate(doc.tables):
        entry = {"index": i, "rows": len(table.rows), "cols": len(table.columns)}
        if include_cells:
            entry["cells"] = [
                [cell.text for cell in row.cells] for row in table.rows
            ]
        tables.append(entry)
    sections = []
    for i, section in enumerate(doc.sections):
        sections.append(
            {
                "index": i,
                "header_text": "\n".join(p.text for p in section.header.paragraphs).strip(),
                "footer_text": "\n".join(p.text for p in section.footer.paragraphs).strip(),
            }
        )
    return {"doc_type": "word", "paragraphs": paragraphs, "tables": tables, "sections": sections}


def _excel_structure(session) -> dict:
    wb = session.doc
    sheets = []
    for name in wb.sheetnames:
        ws = wb[name]
        preview_rows = min(ws.max_row, 20)
        preview_cols = min(ws.max_column, 10)
        preview = [
            [ws.cell(row=r, column=c).value for c in range(1, preview_cols + 1)]
            for r in range(1, preview_rows + 1)
        ]
        sheets.append(
            {
                "name": name,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "preview": preview,
                "preview_note": (
                    f"first {preview_rows}x{preview_cols} cells; use excel_read_cells "
                    "for the rest"
                ),
                "charts": len(getattr(ws, "_charts", [])),
            }
        )
    return {"doc_type": "excel", "sheets": sheets}


@REGISTRY.register(
    "open_document",
    "Open a Word (.docx) or Excel (.xlsx/.xlsm) document for editing. Creates an "
    "isolated working copy — the original file is never modified until save_document "
    "with explicit approval. Returns doc_id used by all other tools.",
    OpenDocumentInput,
)
def open_document(ctx: ToolContext, p: OpenDocumentInput) -> dict:
    session = ctx.sessions.open(p.file_path)
    ctx.snapshot_manager(session.doc_id)  # create baseline snapshot
    if session.doc_type == "word":
        summary = {
            "paragraphs": len(session.doc.paragraphs),
            "tables": len(session.doc.tables),
            "sections": len(session.doc.sections),
        }
    else:
        summary = {"sheets": session.doc.sheetnames}
    return {
        "doc_id": session.doc_id,
        "doc_type": session.doc_type,
        "original_path": str(session.original_path),
        "summary": summary,
    }


@REGISTRY.register(
    "get_structure",
    "Get the full structure of an open document. Word: every paragraph with index, "
    "full text, style name; tables with cell contents; section headers/footers. "
    "Excel: sheets with dimensions and a value preview. Use this before editing to "
    "find correct indices/ranges.",
    GetStructureInput,
)
def get_structure(ctx: ToolContext, p: GetStructureInput) -> dict:
    session = ctx.sessions.get(p.doc_id)
    if session.doc_type == "word":
        return _word_structure(session, p.max_paragraph_chars, p.include_table_cells)
    return _excel_structure(session)


@REGISTRY.register(
    "save_document",
    "Persist the edited document. Without a path, saves to <name>.edited.<ext> next "
    "to the original. Overwriting any pre-existing file requires overwrite=true, "
    "which must only be used after explicit user approval.",
    SaveDocumentInput,
    mutates=False,
)
def save_document(ctx: ToolContext, p: SaveDocumentInput) -> dict:
    session = ctx.sessions.get(p.doc_id)
    dest = p.path if p.path else str(session.default_output_path())
    saved = session.save_to(dest, overwrite=p.overwrite)
    return {"saved_path": str(saved)}


@REGISTRY.register(
    "undo",
    "Undo the last N edit steps (each mutating tool call is one step).",
    UndoInput,
    mutates=False,
)
def undo(ctx: ToolContext, p: UndoInput) -> dict:
    return ctx.snapshot_manager(p.doc_id).undo(p.steps)


@REGISTRY.register(
    "list_snapshots",
    "List edit-step snapshots for a document (for undo / restore_snapshot).",
    ListSnapshotsInput,
)
def list_snapshots(ctx: ToolContext, p: ListSnapshotsInput) -> dict:
    return {"snapshots": ctx.snapshot_manager(p.doc_id).list()}


@REGISTRY.register(
    "restore_snapshot",
    "Restore the document to a specific snapshot from list_snapshots.",
    RestoreSnapshotInput,
    mutates=False,
)
def restore_snapshot(ctx: ToolContext, p: RestoreSnapshotInput) -> dict:
    return ctx.snapshot_manager(p.doc_id).restore(p.snapshot_id)
