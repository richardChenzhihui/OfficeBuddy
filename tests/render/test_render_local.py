"""Local integration tests for the render pipeline.

Require Microsoft Word/Excel on this Mac. Run explicitly:
    pytest -m mac_office
They verify the zero-dialog invariant indirectly: a blocking dialog would
exceed the subprocess timeout and fail the test.
"""
import pytest

from office_agent.core.session import EditSession
from office_agent.render import Renderer, diff_pages

pytestmark = pytest.mark.mac_office


@pytest.fixture
def _container_work_root(monkeypatch):
    """Render tests MUST place working copies inside the app sandbox
    containers (zero-dialog invariant) — undo conftest's unit-test override."""
    monkeypatch.delenv("OFFICE_AGENT_WORK_ROOT", raising=False)


@pytest.fixture
def word_session(word_doc_path, _container_work_root):
    session = EditSession(str(word_doc_path))
    yield session
    session.cleanup()


@pytest.fixture
def excel_session(excel_doc_path, _container_work_root):
    session = EditSession(str(excel_doc_path))
    yield session
    session.cleanup()


def test_word_render_produces_pages(word_session):
    renderer = Renderer(word_session)
    images = renderer.render()
    assert images, "no pages rendered"
    assert images[0].path.exists()
    assert images[0].width > 500


def test_word_render_cache_hit_is_free(word_session):
    renderer = Renderer(word_session)
    first = renderer.render()
    import time

    t0 = time.time()
    second = renderer.render()
    assert time.time() - t0 < 0.5  # cache hit: no Word round-trip
    assert [i.path for i in first] == [i.path for i in second]


def test_word_diff_detects_change_and_bbox(word_session):
    renderer = Renderer(word_session)
    before = renderer.render()
    word_session.doc.paragraphs[0].text = "COMPLETELY DIFFERENT HEADING"
    after = renderer.render()
    diffs = diff_pages(before, after)
    changed = [d for d in diffs if d.changed]
    assert changed, "edit was not detected in rendered output"
    assert changed[0].bbox is not None


def test_word_locate_text_finds_page(word_session):
    renderer = Renderer(word_session)
    renderer.render()
    snippet = word_session.doc.paragraphs[0].text
    if snippet.strip():
        pages = renderer.locate_text(snippet)
        assert 0 in pages


def test_excel_render_produces_pages(excel_session):
    renderer = Renderer(excel_session)
    images = renderer.render()
    assert images
    assert images[0].path.exists()
