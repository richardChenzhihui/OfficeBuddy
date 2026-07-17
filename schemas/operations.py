"""Operation schemas for document editing"""
from typing import Optional, Dict, Any, List, Union
from pydantic import BaseModel, Field
from enum import Enum


class TextOperation(str, Enum):
    """Text editing operations"""
    REPLACE = "replace"
    INSERT = "insert"
    DELETE = "delete"
    APPEND = "append"


class FillMode(str, Enum):
    """Excel fill modes"""
    OVERWRITE = "overwrite"
    APPEND = "append"
    MERGE = "merge"


class StyleParams(BaseModel):
    """Style parameters for formatting
    
    Word and Excel common parameters:
    - font_name: Font family name (e.g., "Times New Roman", "Arial", "Calibri")
    - font_size: Font size in points (float, e.g., 12.0, 14.0)
    - bold: Bold formatting (bool)
    - italic: Italic formatting (bool)
    - underline: Underline formatting (bool)
    - color: Text color in hex format (e.g., "#FF0000" for red)
    - bg_color: Background color in hex format
    - alignment: Text alignment ("left", "center", "right", "justify" for Word)
    
    Excel specific:
    - number_format: Number format string (e.g., "#,##0", "0.00%")
    - border: Border settings dictionary
    
    Word specific:
    - paragraph_spacing: Space before paragraph in points
    - line_spacing: Line spacing multiplier
    """
    font_name: Optional[str] = Field(None, description="Font name (e.g., 'Times New Roman', 'Arial', 'Calibri')")
    font_size: Optional[float] = Field(None, description="Font size in points (e.g., 12.0, 14.0, 16.0)")
    bold: Optional[bool] = Field(None, description="Bold text (True/False)")
    italic: Optional[bool] = Field(None, description="Italic text (True/False)")
    underline: Optional[bool] = Field(None, description="Underline text (True/False)")
    color: Optional[str] = Field(None, description="Text color in hex format (e.g., '#FF0000' for red) or RGB")
    bg_color: Optional[str] = Field(None, description="Background color in hex format")
    alignment: Optional[str] = Field(None, description="Text alignment: 'left', 'center', 'right', 'justify' (Word only)")
    # Excel specific
    number_format: Optional[str] = Field(None, description="Number format string (e.g., '#,##0', '0.00%', 'mm/dd/yyyy')")
    border: Optional[Dict[str, Any]] = Field(None, description="Border settings dictionary")
    # Word specific
    paragraph_spacing: Optional[float] = Field(None, description="Space before paragraph in points")
    line_spacing: Optional[float] = Field(None, description="Line spacing multiplier")


class WordEditOperation(BaseModel):
    """Word editing operation"""
    operation: TextOperation
    selector: Dict[str, Any]  # WordSelector dict
    text: Optional[str] = None
    style: Optional[StyleParams] = None
    element_type: Optional[str] = None  # For insert: table, image, page_break
    content: Optional[Any] = None  # Content for inserted elements


class ExcelEditOperation(BaseModel):
    """Excel editing operation"""
    operation: str = Field(..., description="Operation: write, formula, style, insert, chart")
    selector: Dict[str, Any]  # ExcelSelector dict
    values: Optional[Union[List[List[Any]], Dict[str, Any]]] = None
    formula: Optional[str] = None
    style: Optional[StyleParams] = None
    fill_mode: Optional[FillMode] = FillMode.OVERWRITE
    chart_type: Optional[str] = None
    chart_options: Optional[Dict[str, Any]] = None


class PreviewResult(BaseModel):
    """Preview result for changes"""
    preview_id: str
    diff: str
    affected: List[str]
    summary: Dict[str, Any]


class ChangeResult(BaseModel):
    """Result of applying changes"""
    success: bool
    snapshot_id: Optional[str] = None
    error: Optional[str] = None
    affected_count: Optional[int] = None
