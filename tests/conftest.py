import shutil
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent


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
