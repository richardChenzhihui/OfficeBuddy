"""Selector schemas for document element selection."""
from typing import Any, Dict, Literal, Optional, Tuple, Union

from pydantic import BaseModel, Field


class WordSelector(BaseModel):
    """Targets elements in a Word document.

    Selector types:
    - paragraph: by index / range, or all paragraphs
    - table: by table_index (+ row_index, cell_index for drilling in)
    - run: runs of paragraph `index` (+ run_index for a single run)
    - header / footer: paragraphs of the section's header/footer
    - text_match: paragraphs whose text matches `contains` or `regex`
    - style_match: paragraphs whose style name equals `style_name`
    """

    type: Literal[
        "paragraph", "table", "run", "header", "footer", "text_match", "style_match"
    ] = Field(..., description="Selector type")
    index: Optional[int] = Field(None, description="Element index, e.g. paragraph 3")
    range: Optional[Tuple[int, int]] = Field(
        None, description="Inclusive index range (start, end)"
    )
    table_index: Optional[int] = Field(None, description="Table index")
    row_index: Optional[int] = Field(None, description="Row index within table")
    cell_index: Optional[int] = Field(None, description="Cell index within row")
    run_index: Optional[int] = Field(None, description="Run index within paragraph")
    section_index: int = Field(
        0, description="Section index for header/footer selectors"
    )
    # text_match fields
    contains: Optional[str] = Field(
        None, description="Substring the paragraph text must contain (text_match)"
    )
    regex: Optional[str] = Field(
        None, description="Regex the paragraph text must match (text_match)"
    )
    case_sensitive: bool = Field(True, description="Case sensitivity for text_match")
    occurrence: Union[Literal["first", "all"], int] = Field(
        "first",
        description="Which match to select: 'first', 'all', or a 0-based match index",
    )
    # style_match fields
    style_name: Optional[str] = Field(
        None, description="Paragraph style name to match (style_match), e.g. 'Heading 1'"
    )


class ExcelSelector(BaseModel):
    """Targets cells in an Excel workbook."""

    sheet: str = Field(..., description="Sheet name")
    range: Optional[str] = Field(None, description="A1-style range, e.g. A1:B10 or A:A")
    condition: Optional[Dict[str, Any]] = Field(
        None, description="Condition dict for conditional selection"
    )

    def parse_range(self) -> Optional[Tuple[int, int, Optional[int], Optional[int]]]:
        """Return (min_row, min_col, max_row, max_col) or None for whole sheet.

        Raises ValueError with an actionable message on an unparsable range.
        """
        if not self.range:
            return None
        from openpyxl.utils import range_boundaries

        try:
            min_col, min_row, max_col, max_row = range_boundaries(self.range)
        except Exception as exc:
            raise ValueError(
                f"Invalid Excel range '{self.range}': {exc}. "
                "Use A1-style ranges like 'A1', 'A1:B10', or 'A:A'."
            ) from exc
        # Whole-column ('A:A') / whole-row ('1:1') ranges yield None mins from
        # openpyxl — normalize so downstream arithmetic never sees None.
        return (min_row or 1, min_col or 1, max_row, max_col)
