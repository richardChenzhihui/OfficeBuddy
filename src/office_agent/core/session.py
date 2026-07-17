"""Edit sessions: working-copy isolation inside the Office app sandbox containers.

The original file is never written until an explicit, gated save. The working
copy lives inside the target app's sandbox container so that AppleScript-driven
rendering never triggers macOS file-access permission dialogs.
"""
import json
import shutil
import time
import uuid
from pathlib import Path
from typing import Any, Dict, Optional

from docx import Document as WordDocument
from openpyxl import load_workbook

WORD_CONTAINER = Path.home() / "Library/Containers/com.microsoft.Word/Data/tmp"
EXCEL_CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/tmp"
FALLBACK_ROOT = Path.home() / ".office_agent/sessions"

_EXT_TO_TYPE = {".docx": "word", ".xlsx": "excel", ".xlsm": "excel"}


def work_root_for(doc_type: str) -> Path:
    """Working dir root: the app's sandbox container (prompt-free rendering)."""
    container = WORD_CONTAINER if doc_type == "word" else EXCEL_CONTAINER
    if container.parent.exists():  # app container present
        return container / "office_agent"
    return FALLBACK_ROOT


class EditSession:
    """One open document: working copy, in-memory object, audit log."""

    def __init__(self, file_path: str):
        original = Path(file_path).expanduser().resolve()
        if not original.exists():
            raise FileNotFoundError(f"Document not found: {original}")
        ext = original.suffix.lower()
        if ext not in _EXT_TO_TYPE:
            raise ValueError(
                f"Unsupported file type '{ext}': expected one of {sorted(_EXT_TO_TYPE)}."
            )
        self.doc_type = _EXT_TO_TYPE[ext]
        self.original_path = original
        self.doc_id = uuid.uuid4().hex[:8]
        self.session_dir = work_root_for(self.doc_type) / self.doc_id
        self.session_dir.mkdir(parents=True, exist_ok=True)
        self.working_path = self.session_dir / f"working{ext}"
        shutil.copy2(original, self.working_path)
        self.doc: Any = self._load()
        self.audit_path = self.session_dir / "audit.jsonl"
        self.dirty = False  # in-memory object differs from working_path on disk

    def _load(self) -> Any:
        if self.doc_type == "word":
            return WordDocument(str(self.working_path))
        return load_workbook(str(self.working_path))

    def flush(self) -> Path:
        """Write the in-memory object to the working copy on disk."""
        self.doc.save(str(self.working_path))
        self.dirty = False
        return self.working_path

    def reload_from_bytes(self, data: bytes) -> None:
        """Replace in-memory object (and working copy) from snapshot bytes."""
        self.working_path.write_bytes(data)
        self.doc = self._load()
        self.dirty = False

    def to_bytes(self) -> bytes:
        import io

        buf = io.BytesIO()
        self.doc.save(buf)
        return buf.getvalue()

    def save_to(self, dest: Optional[str] = None, overwrite: bool = False) -> Path:
        """Persist to a real destination. Overwriting the original requires overwrite=True."""
        target = Path(dest).expanduser().resolve() if dest else self.original_path
        if target == self.original_path and not overwrite:
            raise PermissionError(
                f"Refusing to overwrite the original file {target} without explicit "
                "approval. Save to a new path, or pass overwrite=True after user consent."
            )
        self.flush()
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.working_path, target)
        return target

    def default_output_path(self) -> Path:
        p = self.original_path
        return p.with_name(f"{p.stem}.edited{p.suffix}")

    def audit(self, event: str, detail: Dict[str, Any]) -> None:
        record = {"ts": time.time(), "event": event, **detail}
        with self.audit_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")

    def cleanup(self) -> None:
        shutil.rmtree(self.session_dir, ignore_errors=True)


class SessionManager:
    """Registry of open documents for one agent process."""

    def __init__(self) -> None:
        self.sessions: Dict[str, EditSession] = {}

    def open(self, file_path: str) -> EditSession:
        session = EditSession(file_path)
        self.sessions[session.doc_id] = session
        return session

    def get(self, doc_id: str) -> EditSession:
        if doc_id not in self.sessions:
            raise KeyError(
                f"Unknown doc_id '{doc_id}': open documents are {list(self.sessions)}. "
                "Call open_document first."
            )
        return self.sessions[doc_id]

    def close_all(self, keep_dirs: bool = False) -> None:
        for session in self.sessions.values():
            if not keep_dirs:
                session.cleanup()
        self.sessions.clear()
