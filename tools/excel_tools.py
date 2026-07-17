"""Excel editing tools"""
import uuid
from typing import Dict, Any, Optional, List
try:
    from langchain.tools import tool
except ImportError:
    from langchain_core.tools import tool
from pydantic import BaseModel, Field

from .base_tools import get_document_manager
from ..core.selector_parser import SelectorParser
from ..adapters.excel_adapter import ExcelAdapter
from ..schemas.selector import ExcelSelector
from ..schemas.operations import StyleParams, FillMode


class ExcelReadCellsInput(BaseModel):
    """Input for excel_read_cells tool"""
    doc_id: str = Field(..., description="Document ID")
    sheet: str = Field(..., description="Sheet name")
    range: Optional[str] = Field(None, description="Cell range (e.g., A1:B10, A:A)")
    include_formula: bool = Field(False, description="Include formulas in output")


class ExcelWriteCellsInput(BaseModel):
    """Input for excel_write_cells tool"""
    doc_id: str = Field(..., description="Document ID")
    sheet: str = Field(..., description="Sheet name")
    range: Optional[str] = Field(None, description="Cell range (e.g., A1:B10)")
    values: List[List[Any]] = Field(..., description="2D array of values to write")
    fill_mode: str = Field("overwrite", description="Fill mode: overwrite, append, merge")


class ExcelEditFormulaInput(BaseModel):
    """Input for excel_edit_formula tool"""
    doc_id: str = Field(..., description="Document ID")
    sheet: str = Field(..., description="Sheet name")
    range: str = Field(..., description="Cell range (e.g., A1:B10)")
    formula: str = Field(..., description="Formula to set (e.g., =SUM(A1:A10))")


class ExcelEditStyleInput(BaseModel):
    """Input for excel_edit_style tool"""
    doc_id: str = Field(..., description="Document ID")
    sheet: str = Field(..., description="Sheet name")
    range: str = Field(..., description="Cell range")
    style_params: Dict[str, Any] = Field(..., description="Style parameters")


class ExcelInsertRowsColsInput(BaseModel):
    """Input for excel_insert_rows_cols tool"""
    doc_id: str = Field(..., description="Document ID")
    sheet: str = Field(..., description="Sheet name")
    position: int = Field(..., description="Insert position (row or column number)")
    count: int = Field(1, description="Number of rows/columns to insert")
    insert_type: str = Field("row", description="Type: row or column")


class ExcelCreateChartInput(BaseModel):
    """Input for excel_create_chart tool"""
    doc_id: str = Field(..., description="Document ID")
    sheet: str = Field(..., description="Sheet name")
    data_range: str = Field(..., description="Data range for chart (e.g., A1:B10)")
    chart_type: str = Field(..., description="Chart type: bar, line, pie")
    chart_options: Optional[Dict[str, Any]] = Field(None, description="Chart options")


class ExcelConditionalSelectInput(BaseModel):
    """Input for excel_conditional_select tool"""
    doc_id: str = Field(..., description="Document ID")
    sheet: str = Field(..., description="Sheet name")
    condition: Dict[str, Any] = Field(..., description="Condition dictionary")


_doc_manager = get_document_manager()
_selector_parser = SelectorParser()


def _create_preview(doc_id: str, operation: str, result: Dict[str, Any], 
                   extra_data: Optional[Dict[str, Any]] = None) -> str:
    """Create preview and return preview_id"""
    preview_id = str(uuid.uuid4())
    preview_data = {
        "operation": operation,
        "result": result
    }
    if extra_data:
        preview_data.update(extra_data)
    _doc_manager.create_preview(doc_id, preview_id, preview_data)
    return preview_id


@tool(args_schema=ExcelReadCellsInput)
def excel_read_cells(doc_id: str, sheet: str, range: Optional[str] = None, 
                     include_formula: bool = False) -> Dict[str, Any]:
    """Read cells from an Excel worksheet.
    
    Args:
        doc_id: Document ID
        sheet: Sheet name
        range: Optional cell range (e.g., A1:B10, A:A)
        include_formula: Whether to include formulas in output
    
    Returns:
        Dictionary with data and optionally formulas
    """
    workbook = _doc_manager.get_document(doc_id)
    selector = ExcelSelector(sheet=sheet, range=range)
    worksheet, coords = _selector_parser.parse_excel_selector(selector, workbook)
    
    if worksheet is None:
        return {"error": f"Sheet '{sheet}' not found"}
    
    result = ExcelAdapter.read_cells(worksheet, coords, include_formula)
    return result


@tool(args_schema=ExcelWriteCellsInput)
def excel_write_cells(doc_id: str, sheet: str, range: Optional[str], 
                      values: List[List[Any]], fill_mode: str = "overwrite") -> Dict[str, Any]:
    """Write cells to an Excel worksheet (preview mode).
    
    Args:
        doc_id: Document ID
        sheet: Sheet name
        range: Optional cell range (e.g., A1:B10)
        values: 2D array of values to write
        fill_mode: Fill mode (overwrite, append, merge)
    
    Returns:
        Dictionary with preview_id and preview information
    """
    workbook = _doc_manager.get_document(doc_id)
    selector = ExcelSelector(sheet=sheet, range=range)
    worksheet, coords = _selector_parser.parse_excel_selector(selector, workbook)
    
    if worksheet is None:
        return {"error": f"Sheet '{sheet}' not found"}
    
    fill_mode_enum = FillMode(fill_mode)
    result = ExcelAdapter.write_cells(worksheet, coords, values, fill_mode_enum)
    
    preview_id = _create_preview(doc_id, "excel_write_cells", result, {
        "sheet": sheet,
        "range": range,
        "values": values,
        "fill_mode": fill_mode,
        "selector": {"sheet": sheet, "range": range}
    })
    
    return {
        "preview_id": preview_id,
        "preview": result,
        "diff": f"Will write {len(values)} rows to {result.get('range', 'range')}"
    }


@tool(args_schema=ExcelEditFormulaInput)
def excel_edit_formula(doc_id: str, sheet: str, range: str, formula: str) -> Dict[str, Any]:
    """Edit formula in Excel cells (preview mode).
    
    Args:
        doc_id: Document ID
        sheet: Sheet name
        range: Cell range (e.g., A1:B10)
        formula: Formula to set
    
    Returns:
        Dictionary with preview_id and preview information
    """
    workbook = _doc_manager.get_document(doc_id)
    selector = ExcelSelector(sheet=sheet, range=range)
    worksheet, coords = _selector_parser.parse_excel_selector(selector, workbook)
    
    if worksheet is None:
        return {"error": f"Sheet '{sheet}' not found"}
    
    result = ExcelAdapter.edit_formula(worksheet, coords, formula)
    
    preview_id = _create_preview(doc_id, "excel_edit_formula", result, {
        "sheet": sheet,
        "range": range,
        "formula": formula,
        "selector": {"sheet": sheet, "range": range}
    })
    
    return {
        "preview_id": preview_id,
        "preview": result,
        "diff": f"Will set formula '{formula}' in {len(result.get('affected', []))} cells"
    }


@tool(args_schema=ExcelEditStyleInput)
def excel_edit_style(doc_id: str, sheet: str, range: str, 
                     style_params: Dict[str, Any]) -> Dict[str, Any]:
    """Edit style of Excel cells (preview mode).
    
    Available style parameters in style_params:
    - font_name: str - Font name (e.g., "Times New Roman", "Arial", "Calibri")
    - font_size: float - Font size in points (e.g., 12.0, 14.0, 16.0)
    - bold: bool - Bold text (True/False)
    - italic: bool - Italic text (True/False)
    - underline: bool - Underline text (True/False)
    - color: str - Text color in hex format (e.g., "#FF0000" for red)
    - bg_color: str - Background/fill color in hex format
    - alignment: str - Cell alignment: "left", "center", "right"
    - number_format: str - Number format (e.g., "#,##0", "0.00%", "mm/dd/yyyy")
    
    Example style_params:
    {
        "font_name": "Arial",
        "font_size": 14.0,
        "bold": True,
        "italic": False,
        "underline": True,
        "bg_color": "#FFFF00"
    }
    
    Args:
        doc_id: Document ID
        sheet: Sheet name
        range: Cell range (e.g., "A1:B10" or "A1")
        style_params: Style parameters dictionary with any of the above parameters
    
    Returns:
        Dictionary with preview_id and preview information
    """
    workbook = _doc_manager.get_document(doc_id)
    selector = ExcelSelector(sheet=sheet, range=range)
    worksheet, coords = _selector_parser.parse_excel_selector(selector, workbook)
    
    if worksheet is None:
        return {"error": f"Sheet '{sheet}' not found"}
    
    style = StyleParams(**style_params)
    result = ExcelAdapter.apply_style(worksheet, coords, style)
    
    preview_id = _create_preview(doc_id, "excel_edit_style", result, {
        "sheet": sheet,
        "range": range,
        "style_params": style_params,
        "selector": {"sheet": sheet, "range": range}
    })
    
    return {
        "preview_id": preview_id,
        "preview": result,
        "diff": f"Will apply style to {len(result.get('affected', []))} cells"
    }


@tool(args_schema=ExcelInsertRowsColsInput)
def excel_insert_rows_cols(doc_id: str, sheet: str, position: int, 
                           count: int = 1, insert_type: str = "row") -> Dict[str, Any]:
    """Insert rows or columns in Excel (preview mode).
    
    Args:
        doc_id: Document ID
        sheet: Sheet name
        position: Insert position
        count: Number of rows/columns to insert
        insert_type: Type: row or column
    
    Returns:
        Dictionary with preview_id and preview information
    """
    workbook = _doc_manager.get_document(doc_id)
    if sheet not in workbook.sheetnames:
        return {"error": f"Sheet '{sheet}' not found"}
    
    worksheet = workbook[sheet]
    result = ExcelAdapter.insert_rows_cols(worksheet, position, count, insert_type)
    
    preview_id = _create_preview(doc_id, "excel_insert_rows_cols", result, {
        "sheet": sheet,
        "position": position,
        "count": count,
        "insert_type": insert_type
    })
    
    return {
        "preview_id": preview_id,
        "preview": result,
        "diff": f"Will insert {count} {insert_type}(s) at position {position}"
    }


@tool(args_schema=ExcelCreateChartInput)
def excel_create_chart(doc_id: str, sheet: str, data_range: str, 
                       chart_type: str, chart_options: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Create chart in Excel (preview mode).
    
    Args:
        doc_id: Document ID
        sheet: Sheet name
        data_range: Data range for chart
        chart_type: Chart type (bar, line, pie)
        chart_options: Optional chart options
    
    Returns:
        Dictionary with preview_id and preview information
    """
    workbook = _doc_manager.get_document(doc_id)
    if sheet not in workbook.sheetnames:
        return {"error": f"Sheet '{sheet}' not found"}
    
    worksheet = workbook[sheet]
    result = ExcelAdapter.create_chart(worksheet, data_range, chart_type, chart_options)
    
    if "error" in result:
        return result
    
    preview_id = _create_preview(doc_id, "excel_create_chart", result, {
        "sheet": sheet,
        "data_range": data_range,
        "chart_type": chart_type,
        "chart_options": chart_options
    })
    
    return {
        "preview_id": preview_id,
        "preview": result,
        "diff": f"Will create {chart_type} chart from {data_range}"
    }


@tool(args_schema=ExcelConditionalSelectInput)
def excel_conditional_select(doc_id: str, sheet: str, condition: Dict[str, Any]) -> Dict[str, Any]:
    """Select cells based on condition.
    
    Args:
        doc_id: Document ID
        sheet: Sheet name
        condition: Condition dictionary (contains, value_gt, has_formula)
    
    Returns:
        Dictionary with matching cell ranges
    """
    workbook = _doc_manager.get_document(doc_id)
    if sheet not in workbook.sheetnames:
        return {"error": f"Sheet '{sheet}' not found"}
    
    worksheet = workbook[sheet]
    matches = ExcelAdapter.conditional_select(worksheet, condition)
    
    return {
        "matches": matches,
        "count": len(matches)
    }
