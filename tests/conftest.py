import os
import shutil
import tempfile
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent

# Unit tests must never touch the Office app sandbox containers: on machines
# without the TCC container-access grant the syscalls HANG (not error). Render
# tests (mac_office marker) override this per-test where container placement
# actually matters.
os.environ.setdefault(
    "OFFICE_AGENT_WORK_ROOT",
    str(Path(tempfile.gettempdir()) / "office_agent_test_work"),
)


@pytest.fixture
def word_doc_path(tmp_path):
    src = REPO_ROOT / "test.docx"
    dst = tmp_path / "test.docx"
    shutil.copy2(src, dst)
    return dst


@pytest.fixture
def excel_doc_path(tmp_path):
    src = REPO_ROOT / "test.xlsx"
    dst = tmp_path / "test.xlsx"
    shutil.copy2(src, dst)
    return dst
