"""Selector schemas for document element selection"""
from typing import Union, Optional, Dict, Any
from pydantic import BaseModel, Field


class WordSelector(BaseModel):
    """Word document selector"""
    type: str = Field(..., description="Selector type: paragraph, table, run, header, footer")
    index: Optional[int] = Field(None, description="Element index (e.g., paragraph[3])")
    range: Optional[tuple[int, int]] = Field(None, description="Range selector (start, end)")
    table_index: Optional[int] = Field(None, description="Table index for table.selectors")
    row_index: Optional[int] = Field(None, description="Row index for table.selectors")
    cell_index: Optional[int] = Field(None, description="Cell index for table.selectors")
    run_index: Optional[int] = Field(None, description="Run index for paragraph.selectors")
    header_footer_type: Optional[str] = Field(None, description="header or footer for header/footer selectors")


class ExcelSelector(BaseModel):
    """Excel document selector"""
    sheet: str = Field(..., description="Sheet name")
    range: Optional[str] = Field(None, description="Cell range (e.g., A1:B10, A:A)")
    condition: Optional[Dict[str, Any]] = Field(None, description="Condition for conditional selection")
    
    def parse_range(self) -> tuple[str, Optional[tuple[int, int, int, int]]]:
        """Parse range string to coordinates (min_row, min_col, max_row, max_col)"""
        if not self.range:
            return self.sheet, None
        
        # Handle column ranges like A:A
        if ':' in self.range and len(self.range.split(':')) == 2:
            start, end = self.range.split(':')
            if start == end and len(start) == 1:
                # Single column range
                col = ord(start.upper()) - ord('A') + 1
                return self.sheet, (1, col, None, col)
        
        # Handle cell ranges like A1:B10
        from openpyxl.utils import range_boundaries
        try:
            min_col, min_row, max_col, max_row = range_boundaries(self.range)
            return self.sheet, (min_row, min_col, max_row, max_col)
        except:
            return self.sheet, None


class Selector(BaseModel):
    """Unified selector for both Word and Excel"""
    doc_type: str = Field(..., description="Document type: word or excel")
    word: Optional[WordSelector] = None
    excel: Optional[ExcelSelector] = None
    
    @classmethod
    def word_selector(cls, **kwargs) -> "Selector":
        """Create Word selector"""
        return cls(doc_type="word", word=WordSelector(**kwargs))
    
    @classmethod
    def excel_selector(cls, sheet: str, range: Optional[str] = None, condition: Optional[Dict] = None) -> "Selector":
        """Create Excel selector"""
        return cls(doc_type="excel", excel=ExcelSelector(sheet=sheet, range=range, condition=condition))
