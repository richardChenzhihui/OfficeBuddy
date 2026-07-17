"""Excel editing tools. All mutations apply directly and are snapshotted."""
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field

from ..adapters.excel_adapter import ExcelAdapter
from ..core.selector_parser import SelectorParser
from ..schemas.operations import FillMode, StyleParams
from .registry import REGISTRY, ToolContext


def _excel_session(ctx: ToolContext, doc_id: str):
    session = ctx.sessions.get(doc_id)
    if session.doc_type != "excel":
        raise ValueError(
            f"doc_id '{doc_id}' is a Word document; use word_* tools instead."
        )
    return session


def _sheet_and_coords(session, sheet: str, cell_range: Optional[str]):
    selector = {"sheet": sheet}
    if cell_range:
        selector["range"] = cell_range
    return SelectorParser.parse_excel_selector(selector, session.doc)


class ExcelReadInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    sheet: str = Field(..., description="Sheet name")
    range: Optional[str] = Field(
        None, description="A1-style range, e.g. 'A1:B10'. Omit for the whole sheet."
    )
    include_formula: bool = Field(False, description="Also return cell formulas")


class ExcelWriteInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    sheet: str = Field(..., description="Sheet name")
    range: Optional[str] = Field(
        None,
        description="Top-left A1-style anchor (e.g. 'A3'). Omit to write from A1.",
    )
    values: List[List[Any]] = Field(
        ..., description="2D row-major values, e.g. [['Name','Score'],['Bob',85]]"
    )
    fill_mode: FillMode = Field(
        FillMode.OVERWRITE,
        description="overwrite | append (concatenate) | merge (only fill empty cells)",
    )


class ExcelFormulaInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    sheet: str = Field(..., description="Sheet name")
    range: str = Field(..., description="Target cell/range, e.g. 'C2' or 'C2:C10'")
    formula: str = Field(..., description="Formula starting with '=', e.g. '=SUM(B2:B10)'")


class ExcelStyleInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    sheet: str = Field(..., description="Sheet name")
    range: str = Field(..., description="Target range, e.g. 'A1:B1'")
    style_params: StyleParams = Field(
        ...,
        description=(
            "Formatting, e.g. {'bold':true,'bg_color':'#FFFF00','number_format':'0.00%'}"
        ),
    )


class ExcelInsertInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    sheet: str = Field(..., description="Sheet name")
    position: int = Field(..., description="1-based row/column index to insert at")
    count: int = Field(1, description="How many rows/columns to insert")
    insert_type: str = Field("row", description="'row' or 'column'")


class ExcelChartInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    sheet: str = Field(..., description="Sheet name")
    data_range: str = Field(..., description="Chart data range, e.g. 'A1:B10'")
    chart_type: str = Field(..., description="'bar', 'line', or 'pie'")
    chart_options: Optional[Dict[str, Any]] = Field(
        None,
        description=(
            "Options: {'title': str, 'chart_cell': 'E1' (anchor), "
            "'titles_from_data': bool}"
        ),
    )


class ExcelConditionalSelectInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    sheet: str = Field(..., description="Sheet name")
    condition: Dict[str, Any] = Field(
        ...,
        description=(
            "Condition: {'contains': str} and/or {'value_gt': number} and/or "
            "{'has_formula': bool}"
        ),
    )


@REGISTRY.register(
    "excel_read_cells",
    "Read cell values (optionally formulas) from a sheet range.",
    ExcelReadInput,
)
def excel_read_cells(ctx: ToolContext, p: ExcelReadInput) -> dict:
    session = _excel_session(ctx, p.doc_id)
    ws, coords = _sheet_and_coords(session, p.sheet, p.range)
    return ExcelAdapter.read_cells(ws, coords, p.include_formula)


@REGISTRY.register(
    "excel_write_cells",
    "Write a 2D block of values starting at the range anchor. Applies immediately "
    "to the working copy; snapshotted and undoable.",
    ExcelWriteInput,
    mutates=True,
)
def excel_write_cells(ctx: ToolContext, p: ExcelWriteInput) -> dict:
    session = _excel_session(ctx, p.doc_id)
    ws, coords = _sheet_and_coords(session, p.sheet, p.range)
    result = ExcelAdapter.write_cells(ws, coords, p.values, p.fill_mode)
    return {"matched_count": len(result["affected"]), **result}


@REGISTRY.register(
    "excel_edit_formula",
    "Set a formula in the given cell(s). Applies immediately; snapshotted.",
    ExcelFormulaInput,
    mutates=True,
)
def excel_edit_formula(ctx: ToolContext, p: ExcelFormulaInput) -> dict:
    session = _excel_session(ctx, p.doc_id)
    ws, coords = _sheet_and_coords(session, p.sheet, p.range)
    return ExcelAdapter.edit_formula(ws, coords, p.formula)


@REGISTRY.register(
    "excel_edit_style",
    "Apply formatting (font, fill, alignment, number format) to a range. "
    "Applies immediately; snapshotted.",
    ExcelStyleInput,
    mutates=True,
)
def excel_edit_style(ctx: ToolContext, p: ExcelStyleInput) -> dict:
    session = _excel_session(ctx, p.doc_id)
    ws, coords = _sheet_and_coords(session, p.sheet, p.range)
    return ExcelAdapter.apply_style(ws, coords, p.style_params)


@REGISTRY.register(
    "excel_insert_rows_cols",
    "Insert empty rows or columns at a position. Applies immediately; snapshotted.",
    ExcelInsertInput,
    mutates=True,
)
def excel_insert_rows_cols(ctx: ToolContext, p: ExcelInsertInput) -> dict:
    session = _excel_session(ctx, p.doc_id)
    ws, _ = _sheet_and_coords(session, p.sheet, None)
    return ExcelAdapter.insert_rows_cols(ws, p.position, p.count, p.insert_type)


@REGISTRY.register(
    "excel_create_chart",
    "Create a bar/line/pie chart from a data range. Applies immediately; snapshotted.",
    ExcelChartInput,
    mutates=True,
)
def excel_create_chart(ctx: ToolContext, p: ExcelChartInput) -> dict:
    session = _excel_session(ctx, p.doc_id)
    ws, _ = _sheet_and_coords(session, p.sheet, None)
    return ExcelAdapter.create_chart(ws, p.data_range, p.chart_type, p.chart_options)


@REGISTRY.register(
    "excel_conditional_select",
    "Find cells matching a condition (contains / value_gt / has_formula). "
    "Read-only; returns matching cell addresses for use in later edits.",
    ExcelConditionalSelectInput,
)
def excel_conditional_select(ctx: ToolContext, p: ExcelConditionalSelectInput) -> dict:
    session = _excel_session(ctx, p.doc_id)
    ws, _ = _sheet_and_coords(session, p.sheet, None)
    matches = ExcelAdapter.conditional_select(ws, p.condition)
    return {"matches": matches, "matched_count": len(matches)}
