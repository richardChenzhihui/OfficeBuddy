"""bench/BUGS.md OA-6: a failed export must not leave the file open in the app.

A workbook left open makes every later `open` of the same name return the
stale in-memory copy, so one failure cascades into an unrecoverable run of
-50 errors. These tests are offline: osascript is never invoked.
"""
from pathlib import Path

import pytest

from office_agent.render import applescript as A


@pytest.fixture
def spy(monkeypatch):
    calls = []

    def fake_run(script, timeout, app):
        calls.append({"script": script, "timeout": timeout, "app": app})
        if "save as" in script or "save workbook as" in script:
            raise A.RenderError("boom: export failed")
        return ""

    monkeypatch.setattr(A, "run_applescript", fake_run)
    return calls


def test_failed_word_export_closes_the_document(spy, tmp_path):
    with pytest.raises(A.RenderError, match="boom"):
        A.export_docx_to_pdf(tmp_path / "working.docx", tmp_path / "out.pdf")
    cleanup = spy[-1]
    assert cleanup["app"] == "Microsoft Word"
    assert 'if name of d is "working.docx" then close d saving no' in cleanup["script"]
    assert "every document" in cleanup["script"]
    assert "quit" not in cleanup["script"]  # never kill the user's own session


def test_failed_excel_export_closes_the_workbook(spy, tmp_path):
    with pytest.raises(A.RenderError, match="boom"):
        A.export_xlsx_to_pdf(tmp_path / "working.xlsx", tmp_path / "out.pdf")
    cleanup = spy[-1]
    assert cleanup["app"] == "Microsoft Excel"
    assert 'if name of d is "working.xlsx" then close d saving no' in cleanup["script"]
    assert "every workbook" in cleanup["script"]


def test_cleanup_never_launches_a_stopped_app(spy, tmp_path):
    A.close_stale_document("Microsoft Excel", "working.xlsx")
    assert spy[-1]["script"].lstrip().startswith(
        'if application "Microsoft Excel" is running then'
    )


def test_cleanup_failure_does_not_mask_the_real_error(monkeypatch, tmp_path):
    def fake_run(script, timeout, app):
        if "close d saving no" in script:
            raise A.RenderTimeout("cleanup also hung")
        raise A.RenderError("original failure")

    monkeypatch.setattr(A, "run_applescript", fake_run)
    with pytest.raises(A.RenderError, match="original failure"):
        A.export_xlsx_to_pdf(tmp_path / "working.xlsx", tmp_path / "out.pdf")


def test_missing_pdf_also_triggers_cleanup(monkeypatch, tmp_path):
    """Export 'succeeded' but produced nothing — the doc is still open."""
    calls = []
    monkeypatch.setattr(
        A, "run_applescript", lambda script, timeout, app: calls.append(script) or ""
    )
    with pytest.raises(A.RenderError, match="no PDF was produced"):
        A.export_docx_to_pdf(tmp_path / "working.docx", tmp_path / "out.pdf")
    assert "close d saving no" in calls[-1]


def test_cleanup_escapes_the_filename(spy):
    A.close_stale_document("Microsoft Excel", 'we"ird\\name.xlsx')
    assert r'we\"ird\\name.xlsx' in spy[-1]["script"]
