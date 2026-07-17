"""Edit sessions: working-copy isolation inside the Office app sandbox containers.

The original file is never written until an explicit, gated save. The working
copy lives inside the target app's sandbox container so that AppleScript-driven
rendering never triggers macOS file-access permission dialogs.
"""
import io
import json
import os
import shutil
import time
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document as WordDocument
from openpyxl import load_workbook

WORD_CONTAINER = Path.home() / "Library/Containers/com.microsoft.Word/Data/tmp"
EXCEL_CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/tmp"
FALLBACK_ROOT = Path.home() / ".office_agent/sessions"

_EXT_TO_TYPE = {".docx": "word", ".xlsx": "excel", ".xlsm": "excel"}


@dataclass(frozen=True)
class FidelityLossEntry:
    path: str  # zip member name, e.g. "xl/threadedComments/threadedComment1.xml"
    size: int  # bytes, from the ORIGINAL file
    category: str  # human label, e.g. "threaded comment"


class FidelityLossError(PermissionError):
    """Saving would silently drop OOXML parts openpyxl's writer cannot preserve.

    Subclasses PermissionError so it flows through ToolRegistry.dispatch's
    existing caught-exception tuple unchanged — same precedent as the
    overwrite guard in save_to()."""


_FIDELITY_CATEGORIES = [  # longest/most-specific prefix first
    ("xl/threadedComments/", "threaded comment"),
    ("xl/persons/", "comment-author registry"),
    ("xl/slicerCaches/", "slicer cache"),
    ("xl/slicers/", "slicer"),
    ("xl/pivotCache/", "pivot cache"),
    ("xl/pivotTables/", "pivot table"),
    ("xl/richData/", "rich data type (stocks/geography)"),
    ("xl/metadata.xml", "rich-value/dynamic-array metadata"),
    ("customXml/", "custom XML part"),
    ("xl/ctrlProps/", "form control"),
    ("xl/activeX/", "ActiveX object"),
    ("customUI/", "ribbon customization"),
    ("xl/media/", "embedded image/media"),
    ("xl/charts/", "chart"),
    ("xl/drawings/", "drawing (shape/textbox/chart anchor)"),
]


def _categorize(name: str) -> str:
    for prefix, label in _FIDELITY_CATEGORIES:
        if name.startswith(prefix):
            return label
    return "unrecognized part"


def _zip_inventory(source) -> Dict[str, int]:
    """source: Path or bytes. Returns {member_name: size}. Cheap — reads only
    the central directory, no decompression."""
    fh = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
    with zipfile.ZipFile(fh) as zf:
        return {i.filename: i.file_size for i in zf.infolist()}


def _format_fidelity_error(lost: List[FidelityLossEntry], original: Path) -> str:
    lines = "\n".join(f"  - {e.category}: {e.path} ({e.size} bytes)" for e in lost)
    return (
        f"Saving would silently DROP {len(lost)} part(s) that this editing "
        "engine (openpyxl) cannot preserve because it fully re-serializes the "
        f"workbook from its own object model:\n{lines}\n"
        f"The original file at {original} is untouched.\n"
        "Options: (1) show the user exactly what would be lost and, only after "
        "explicit confirmation, retry save_document with "
        "accept_fidelity_loss=True; (2) save to a new path so the user keeps "
        "the original intact for these features."
    )


def work_root_for(doc_type: str) -> tuple:
    """(root, sandboxed): the app's sandbox container when present (prompt-free
    rendering), else a private fallback where rendering WILL prompt.

    OFFICE_AGENT_WORK_ROOT overrides everything (tests/dev): macOS app-container
    protection can block access to other apps' containers for processes without
    the TCC grant — the syscall HANGS rather than erroring — so unit tests must
    never depend on container access."""
    override = os.environ.get("OFFICE_AGENT_WORK_ROOT")
    if override:
        return Path(override) / "sessions", True  # quiet: explicit choice
    container = WORD_CONTAINER if doc_type == "word" else EXCEL_CONTAINER
    if _dir_accessible(container.parent):
        return container / "office_agent", True
    return FALLBACK_ROOT, False


_ACCESS_CACHE: Dict[str, bool] = {}


def _dir_accessible(path: Path, timeout: float = 3.0) -> bool:
    """True if the directory exists AND is actually readable.

    macOS app-container protection makes syscalls HANG (not error) for
    processes without the TCC grant, so the probe runs in a subprocess with a
    hard timeout — an inaccessible container degrades to the fallback root
    instead of freezing the session forever."""
    key = str(path)
    if key not in _ACCESS_CACHE:
        import subprocess

        try:
            proc = subprocess.run(
                ["/bin/ls", key], capture_output=True, timeout=timeout
            )
            _ACCESS_CACHE[key] = proc.returncode == 0
        except subprocess.TimeoutExpired:
            _ACCESS_CACHE[key] = False
    return _ACCESS_CACHE[key]


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
        root, self.sandboxed = work_root_for(self.doc_type)
        if not self.sandboxed:
            import sys

            print(
                f"[office-agent] 警告：{self.doc_type} 应用容器不存在，工作目录退回 "
                f"{FALLBACK_ROOT} —— 渲染时可能出现 macOS 文件访问弹窗。"
                "运行 `office-agent doctor` 检查。",
                file=sys.stderr,
            )
        self.session_dir = root / self.doc_id
        self.session_dir.mkdir(parents=True, exist_ok=True)
        # Session dirs hold full document contents + audit logs: owner-only.
        os.chmod(self.session_dir, 0o700)
        if not self.sandboxed:
            os.chmod(root, 0o700)
        self.working_path = self.session_dir / f"working{ext}"
        shutil.copy2(original, self.working_path)
        self._baseline_inventory: Dict[str, int] = (
            _zip_inventory(self.working_path) if self.doc_type == "excel" else {}
        )
        self.doc: Any = self._load()
        # Loss is fixed at load time: openpyxl's object model either captured
        # a part or it never did. One trial serialize here is the same cost
        # class as the baseline undo snapshot taken right after open.
        self.fidelity_report: List[FidelityLossEntry] = self._compute_fidelity_report()
        self.audit_path = self.session_dir / "audit.jsonl"
        self.dirty = False  # in-memory object differs from working_path on disk
        self.written_paths: set = set()  # files this session has saved to

    def _compute_fidelity_report(self) -> List[FidelityLossEntry]:
        if self.doc_type != "excel":
            return []
        current = _zip_inventory(self.to_bytes())
        return [
            FidelityLossEntry(name, size, _categorize(name))
            for name, size in self._baseline_inventory.items()
            if name not in current
        ]

    def fidelity_warnings(self) -> List[str]:
        """Human-readable strings for open_document's `warnings` field."""
        if not self.fidelity_report:
            return []
        by_cat: Dict[str, List[FidelityLossEntry]] = {}
        for entry in self.fidelity_report:
            by_cat.setdefault(entry.category, []).append(entry)
        parts = ", ".join(f"{len(v)} {k}(s)" for k, v in by_cat.items())
        return [
            "⚠️ This workbook contains part(s) the current editing engine "
            f"CANNOT preserve — {parts} will be LOST if you save: "
            f"{', '.join(e.path for e in self.fidelity_report)}. Inform the "
            "user before saving; save_document will require "
            "accept_fidelity_loss=True after explicit user consent."
        ]

    def _load(self) -> Any:
        if self.doc_type == "word":
            return WordDocument(str(self.working_path))
        return load_workbook(
            str(self.working_path),
            # Always on: rescues vbaProject/ctrlProps/activeX/vmlDrawing parts
            # whenever present; verified zero side effects on plain .xlsx.
            keep_vba=True,
            # Explicit (openpyxl default): external-link formulas round-trip.
            keep_links=True,
            # Without this, intra-cell mixed-run formatting is silently
            # flattened to a plain string at load — irreversible loss.
            rich_text=True,
        )

    def flush(self) -> Path:
        """Write the in-memory object to the working copy on disk."""
        self.doc.save(str(self.working_path))
        self.dirty = False
        return self.working_path

    def _close_doc_handles(self) -> None:
        """keep_vba workbooks hold an open ZipFile on the source; close it
        before discarding the object so GC-time destructors stay silent."""
        archive = getattr(self.doc, "vba_archive", None)
        if archive is not None:
            try:
                archive.close()
            except Exception:
                pass

    def reload_from_bytes(self, data: bytes) -> None:
        """Replace in-memory object (and working copy) from snapshot bytes."""
        self._close_doc_handles()
        self.working_path.write_bytes(data)
        self.doc = self._load()
        self.dirty = False

    def to_bytes(self) -> bytes:
        import io

        buf = io.BytesIO()
        self.doc.save(buf)
        return buf.getvalue()

    def save_to(
        self,
        dest: Optional[str] = None,
        overwrite: bool = False,
        accept_fidelity_loss: bool = False,
    ) -> Path:
        """Persist to a real destination.

        Overwriting ANY pre-existing file (the original or an unrelated one)
        requires overwrite=True — except files this session already wrote,
        which may be re-saved freely. Saving an Excel workbook that would drop
        OOXML parts requires accept_fidelity_loss=True (explicit user consent).
        """
        target = Path(dest).expanduser().resolve() if dest else self.original_path
        if target == self.original_path and not overwrite:
            raise PermissionError(
                f"Refusing to overwrite the original file {target} without explicit "
                "approval. Save to a new path, or pass overwrite=True after user consent."
            )
        if (
            target.exists()
            and str(target) not in self.written_paths
            and target != self.original_path
            and not overwrite
        ):
            raise PermissionError(
                f"Refusing to overwrite the existing file {target}: it was not "
                "created by this session. Choose a new path, or pass "
                "overwrite=True after user consent."
            )
        self.flush()
        if self.doc_type == "excel" and not accept_fidelity_loss:
            current = _zip_inventory(self.working_path)
            lost = [
                FidelityLossEntry(name, size, _categorize(name))
                for name, size in self._baseline_inventory.items()
                if name not in current
            ]
            if lost:
                raise FidelityLossError(_format_fidelity_error(lost, self.original_path))
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.working_path, target)
        self.written_paths.add(str(target))
        return target

    def default_output_path(self) -> Path:
        p = self.original_path
        return p.with_name(f"{p.stem}.edited{p.suffix}")

    def audit(self, event: str, detail: Dict[str, Any]) -> None:
        record = {"ts": time.time(), "event": event, **detail}
        with self.audit_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")

    def cleanup(self) -> None:
        self._close_doc_handles()
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
