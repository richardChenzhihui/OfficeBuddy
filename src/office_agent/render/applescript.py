"""AppleScript-driven PDF export via the real Word/Excel apps.

Zero-dialog invariants (validated empirically):
1. Input/output files live inside the target app's sandbox container
   (~/Library/Containers/com.microsoft.<App>/Data/...) — no file-access prompts.
2. The target PDF is deleted before export — no overwrite-confirmation dialog.
3. display alerts is turned off for the scripted operation.
4. Never `activate` — no focus stealing.
Every script runs under both an AppleScript `with timeout` and a subprocess
timeout, so a stuck dialog can never block forever.
"""
import subprocess
import threading
from pathlib import Path

# All Word/Excel automation is serialized: 'active document'/'active workbook'
# references make concurrent exports race each other.
_AUTOMATION_LOCK = threading.Lock()


class RenderError(RuntimeError):
    """PDF export failed for an unclassified reason."""


class RenderTimeout(RenderError):
    """Export did not finish in time (possible blocked dialog or app hang)."""


class AutomationDenied(RenderError):
    """macOS automation (TCC) permission for controlling the app is missing."""


class DocumentCorrupted(RenderError):
    """The app could not open the file — likely corrupted by an earlier write."""


def _classify(stderr: str, app: str) -> RenderError:
    low = stderr.lower()
    if "-1743" in stderr or "not authorized" in low or "not allowed" in low:
        return AutomationDenied(
            f"macOS automation permission for {app} is missing. One-time fix: "
            "System Settings > Privacy & Security > Automation — allow your "
            f"terminal to control {app}. Then re-run. ({stderr.strip()})"
        )
    if "-1712" in stderr or "timed out" in low:
        return RenderTimeout(
            f"{app} did not respond in time — a dialog may be blocking it. "
            f"({stderr.strip()})"
        )
    if any(k in low for k in ("cannot be opened", "damaged", "corrupt", "无法打开", "损坏")):
        return DocumentCorrupted(
            f"{app} could not open the document — the file may have been "
            f"corrupted by a previous edit. ({stderr.strip()})"
        )
    return RenderError(f"{app} PDF export failed: {stderr.strip()}")


def run_applescript(script: str, timeout: float, app: str) -> str:
    try:
        with _AUTOMATION_LOCK:
            proc = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=timeout,
            )
    except subprocess.TimeoutExpired as exc:
        raise RenderTimeout(
            f"{app} export exceeded {timeout}s — a dialog may be blocking it, or "
            "the app is cold-starting. Retry once; if it persists run "
            "`office-agent doctor`."
        ) from exc
    if proc.returncode != 0:
        raise _classify(proc.stderr or proc.stdout, app)
    return proc.stdout


def _q(value) -> str:
    """Escape ANY string (path, sheet name, …) for embedding inside an
    AppleScript string literal: backslashes first, then quotes."""
    return str(value).replace("\\", "\\\\").replace('"', '\\"')


_CLEANUP_TIMEOUT = 20.0


def close_stale_document(app: str, filename: str) -> None:
    """Best-effort: close a document the failed export may have left open.

    A workbook left open in Excel makes every later `open` of the same name
    return the stale in-memory copy, so subsequent exports fail too — the -50
    avalanche in bench/BUGS.md OA-6. Never raises, never launches the app, and
    never quits it (the user may have their own documents open).
    """
    name = _q(filename)
    if app == "Microsoft Word":
        alerts, collection = "set display alerts to alerts none", "every document"
    else:
        alerts, collection = "set display alerts to false", "every workbook"
    script = f'''
if application "{_q(app)}" is running then
    with timeout of {int(_CLEANUP_TIMEOUT) - 5} seconds
        tell application "{_q(app)}"
            try
                {alerts}
            end try
            repeat with d in (get {collection})
                try
                    if name of d is "{name}" then close d saving no
                end try
            end repeat
        end tell
    end timeout
end if
'''
    try:
        run_applescript(script, _CLEANUP_TIMEOUT, app)
    except Exception:  # noqa: BLE001 — cleanup must never mask the real failure
        pass


def export_docx_to_pdf(docx_path: Path, pdf_path: Path, timeout: float = 120.0) -> Path:
    pdf_path.unlink(missing_ok=True)
    script = f'''
with timeout of {int(timeout) - 5} seconds
    tell application "Microsoft Word"
        set display alerts to alerts none
        open (POSIX file "{_q(docx_path)}")
        set theDoc to active document
        save as theDoc file name "{_q(pdf_path)}" file format format PDF
        close theDoc saving no
    end tell
end timeout
'''
    try:
        run_applescript(script, timeout, "Microsoft Word")
    except RenderError:
        close_stale_document("Microsoft Word", docx_path.name)
        raise
    if not pdf_path.exists():
        close_stale_document("Microsoft Word", docx_path.name)
        raise RenderError(
            f"Word reported success but no PDF was produced at {pdf_path}."
        )
    return pdf_path


def export_xlsx_to_pdf(
    xlsx_path: Path,
    pdf_path: Path,
    sheet: str | None = None,
    timeout: float = 120.0,
) -> Path:
    pdf_path.unlink(missing_ok=True)
    activate_sheet = ""
    if sheet:
        sheet_escaped = _q(sheet)
        activate_sheet = f'''
        try
            activate object worksheet "{sheet_escaped}" of theBook
            set fit to pages wide of page setup object of active sheet to 1
            set zoom of page setup object of active sheet to false
        end try'''
    script = f'''
with timeout of {int(timeout) - 5} seconds
    tell application "Microsoft Excel"
        set display alerts to false
        open "{_q(xlsx_path)}"
        set theBook to active workbook{activate_sheet}
        save workbook as theBook filename "{_q(pdf_path)}" file format PDF file format
        close theBook saving no
    end tell
end timeout
'''
    try:
        run_applescript(script, timeout, "Microsoft Excel")
    except RenderError:
        close_stale_document("Microsoft Excel", xlsx_path.name)
        raise
    if not pdf_path.exists():
        close_stale_document("Microsoft Excel", xlsx_path.name)
        raise RenderError(
            f"Excel reported success but no PDF was produced at {pdf_path}."
        )
    return pdf_path
