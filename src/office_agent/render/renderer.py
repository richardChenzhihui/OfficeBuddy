"""High-level document rendering: session -> PDF -> page images, with caching."""
import hashlib
from pathlib import Path
from typing import List, Optional, Tuple

from ..core.session import EditSession
from .applescript import export_docx_to_pdf, export_xlsx_to_pdf
from .pdf_to_images import PageImage, locate_text_pages, pdf_to_images


class Renderer:
    """Renders a session's current in-memory state to page PNGs.

    Caches by content hash: repeated renders of unchanged content are free.
    Keeps the previous render around so callers can diff before/after.
    """

    def __init__(self, session: EditSession):
        self.session = session
        self.render_dir = session.session_dir / "render"
        self.render_dir.mkdir(exist_ok=True)
        self._cache_hash: Optional[str] = None
        self._cache: Optional[Tuple[Path, List[PageImage]]] = None
        self.previous: Optional[List[PageImage]] = None

    def render(self, sheet: Optional[str] = None, timeout: float = 120.0) -> List[PageImage]:
        self.session.flush()
        content = self.session.working_path.read_bytes()
        key_src = content + (sheet or "").encode()
        key = hashlib.sha256(key_src).hexdigest()[:16]
        if key == self._cache_hash and self._cache is not None:
            return self._cache[1]

        out_dir = self.render_dir / key
        pdf_path = out_dir / "render.pdf"
        out_dir.mkdir(parents=True, exist_ok=True)
        if self.session.doc_type == "word":
            export_docx_to_pdf(self.session.working_path, pdf_path, timeout=timeout)
        else:
            export_xlsx_to_pdf(
                self.session.working_path, pdf_path, sheet=sheet, timeout=timeout
            )
        images = pdf_to_images(pdf_path, out_dir)

        if self._cache is not None:
            self.previous = self._cache[1]
        self._cache_hash = key
        self._cache = (pdf_path, images)
        return images

    @property
    def current_pdf(self) -> Optional[Path]:
        return self._cache[0] if self._cache else None

    def locate_text(self, snippet: str) -> List[int]:
        """0-based page indices where the snippet appears in the latest render."""
        if not self.current_pdf:
            return []
        return locate_text_pages(self.current_pdf, snippet)


def check_render_environment() -> dict:
    """Doctor check: is prompt-free rendering possible? Returns findings."""
    import shutil
    import subprocess

    findings = {}
    for app, container in (
        ("Microsoft Word", "com.microsoft.Word"),
        ("Microsoft Excel", "com.microsoft.Excel"),
    ):
        app_installed = Path(f"/Applications/{app}.app").exists()
        container_ok = (Path.home() / f"Library/Containers/{container}").exists()
        automation = None
        if app_installed:
            proc = subprocess.run(
                ["osascript", "-e", f'tell application "{app}" to get version'],
                capture_output=True,
                text=True,
                timeout=60,
            )
            automation = proc.returncode == 0
        findings[app] = {
            "installed": app_installed,
            "container": container_ok,
            "automation_permitted": automation,
        }
    findings["pymupdf"] = True
    findings["osascript"] = shutil.which("osascript") is not None
    return findings
