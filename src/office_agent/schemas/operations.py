"""Operation schemas for document editing."""
from enum import Enum
from typing import Any, Dict, Optional

from pydantic import BaseModel, Field


class TextOperation(str, Enum):
    REPLACE = "replace"
    INSERT = "insert"  # prepend at element start
    DELETE = "delete"
    APPEND = "append"


class FillMode(str, Enum):
    OVERWRITE = "overwrite"
    APPEND = "append"  # concatenate onto existing cell value
    MERGE = "merge"  # only fill empty cells


class StyleParams(BaseModel):
    """Formatting parameters shared by Word and Excel style tools."""

    font_name: Optional[str] = Field(None, description="Font name, e.g. 'Times New Roman'")
    font_size: Optional[float] = Field(None, description="Font size in points, e.g. 12.0")
    bold: Optional[bool] = Field(None, description="Bold on/off")
    italic: Optional[bool] = Field(None, description="Italic on/off")
    underline: Optional[bool] = Field(None, description="Underline on/off")
    color: Optional[str] = Field(None, description="Text color hex, e.g. '#FF0000'")
    bg_color: Optional[str] = Field(None, description="Background color hex (Excel fill)")
    alignment: Optional[str] = Field(
        None, description="Alignment: left / center / right / justify"
    )
    # Excel specific
    number_format: Optional[str] = Field(
        None, description="Excel number format, e.g. '#,##0', '0.00%'"
    )
    border: Optional[Dict[str, Any]] = Field(None, description="Excel border settings")
    # Word specific
    paragraph_spacing: Optional[float] = Field(
        None, description="Space before paragraph in points (Word)"
    )
    line_spacing: Optional[float] = Field(
        None, description="Line spacing multiplier (Word)"
    )
