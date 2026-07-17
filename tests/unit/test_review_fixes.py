"""Regression tests for defects found by the adversarial review."""
import pytest
from docx import Document
from openpyxl import Workbook

from office_agent.adapters.excel_adapter import ExcelAdapter
from office_agent.adapters.word_adapter import WordAdapter
from office_agent.agent.budget import BudgetTracker
from office_agent.agent.plan import Plan
from office_agent.config import Budgets
from office_agent.core.session import EditSession
from office_agent.schemas.operations import StyleParams
from office_agent.schemas.selector import ExcelSelector


# --- find_replace: replace containing find must not double-apply -------------

def test_find_replace_replace_containing_find():
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run("We sell one cat here.")
    run.bold = True
    result = WordAdapter.find_replace(doc, "cat", "cats")
    assert p.text == "We sell one cats here."
    assert result["count"] == 1
    assert "warning" not in result
    assert p.runs[0].bold is True  # formatting preserved


def test_find_replace_suffix_addition():
    doc = Document()
    doc.add_paragraph("Acme Inc is Inc based.")
    result = WordAdapter.find_replace(doc, "Inc", "Inc.")
    assert doc.paragraphs[0].text == "Acme Inc. is Inc. based."
    assert result["count"] == 2


# --- whole-column / whole-row ranges -----------------------------------------

def test_parse_range_whole_column_normalizes_none():
    coords = ExcelSelector(sheet="S", range="A:A").parse_range()
    assert coords[0] == 1 and coords[1] == 1  # min_row/min_col never None


def test_write_cells_whole_column_range():
    wb = Workbook()
    ws = wb.active
    coords = ExcelSelector(sheet="S", range="B:B").parse_range()
    result = ExcelAdapter.write_cells(ws, coords, [["x"], ["y"]])
    assert ws["B1"].value == "x" and ws["B2"].value == "y"
    assert result["affected"] == ["B1", "B2"]


def test_style_whole_row_range():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "h"
    coords = ExcelSelector(sheet="S", range="1:1").parse_range()
    ExcelAdapter.apply_style(ws, coords, StyleParams(bold=True))
    assert ws["A1"].font.bold is True


# --- chart categories --------------------------------------------------------

def test_chart_first_column_becomes_categories():
    wb = Workbook()
    ws = wb.active
    for row in [["名称", "销量"], ["苹果", 120], ["香蕉", 85]]:
        ws.append(row)
    ExcelAdapter.create_chart(ws, "A1:B3", "bar")
    chart = ws._charts[0]
    assert len(chart.series) == 1  # label column is categories, not a series
    assert chart.series[0].cat is not None


# --- fail-before-mutate ------------------------------------------------------

def test_insert_table_bad_position_leaves_no_orphan():
    doc = Document()
    doc.add_paragraph("only one")
    n_tables = len(doc.tables)
    with pytest.raises(ValueError, match="out of range"):
        WordAdapter.insert_element(doc, 99, "table", [["a"]])
    assert len(doc.tables) == n_tables  # no orphaned table appended


def test_apply_style_invalid_alignment_mutates_nothing():
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("text")
    with pytest.raises(ValueError, match="Invalid alignment"):
        WordAdapter.apply_style([p], StyleParams(alignment="middle", bold=True))
    assert p.runs[0].bold is None  # bold was NOT applied before the failure


def test_insert_table_at_position_inserts_before():
    doc = Document()
    doc.add_paragraph("first")
    doc.add_paragraph("second")
    WordAdapter.insert_element(doc, 1, "table", [["cell"]])
    # Table sits between 'first' and 'second' (before paragraph index 1)
    body = list(doc.element.body)
    tags = [el.tag.split("}")[-1] for el in body]
    assert tags.index("tbl") < tags.index("p") + 2  # table not at document end
    assert doc.tables[0].rows[0].cells[0].text == "cell"


# --- plan active_index -------------------------------------------------------

def test_plan_active_index_tracks_most_recent():
    plan = Plan.from_descriptions(["A", "B"])
    plan.set_status(0, "in_progress")
    assert plan.current_index == 0
    # A fails verification and stays in_progress; model moves to B.
    plan.set_status(1, "in_progress")
    assert plan.current_index == 1  # NOT hijacked by lower-index step A
    plan.set_status(1, "done")
    assert plan.current_index == 0  # falls back to still-in_progress A


# --- budget: end-turn cap and step reset -------------------------------------

def test_end_turn_repair_budget_caps():
    tracker = BudgetTracker(Budgets(max_attempts_per_step=2))
    assert not tracker.end_turn_exhausted()
    tracker.record_end_turn_repair()
    tracker.record_end_turn_repair()
    assert tracker.end_turn_exhausted()


def test_reset_steps_preserves_task_totals():
    tracker = BudgetTracker(Budgets())
    tracker.record_tool_call(0)
    tracker.record_failure(0, "sig")
    tracker.reset_steps()
    assert tracker.total_tool_calls == 1  # task counter persists
    assert tracker.step(0).attempts == 0  # per-step state reset


# --- save protection for arbitrary existing files ----------------------------

def test_save_to_refuses_unrelated_existing_file(word_doc_path, tmp_path):
    victim = tmp_path / "victim.docx"
    victim.write_bytes(word_doc_path.read_bytes())
    session = EditSession(str(word_doc_path))
    try:
        with pytest.raises(PermissionError, match="not.*created by this session"):
            session.save_to(str(victim), overwrite=False)
        # But re-saving a file we wrote ourselves is fine.
        out = tmp_path / "mine.docx"
        session.save_to(str(out))
        session.save_to(str(out))  # second save, no error
    finally:
        session.cleanup()


# --- verifier fails closed ---------------------------------------------------

def test_verifier_fails_closed_on_malformed_response():
    from office_agent.agent.verifier import verify_edit

    class TextOnlyLLM:
        def create(self, **kwargs):
            class Block:
                type = "text"
                text = "I think it looks fine!"

            class Msg:
                content = [Block()]

            return Msg()

    verdict = verify_edit(TextOnlyLLM(), "step", "ops", after_images=[])
    assert verdict.passed is False
    assert verdict.blocking
    assert any("structured verdict" in p["description"] for p in verdict.problems)
