"""Resolves selectors into live python-docx / openpyxl objects.

Raises SelectorError with actionable messages for structural problems
(index out of range, missing sheet). Returns [] only when a content match
(text_match / style_match) legitimately finds nothing — callers should
report the zero-match with context.
"""
import re
from typing import Any, List, Optional, Tuple, Union

from ..schemas.selector import ExcelSelector, WordSelector


class SelectorError(ValueError):
    """A selector could not be resolved against the document."""


def _match_occurrence(matches: List[Any], occurrence: Union[str, int]) -> List[Any]:
    if occurrence == "all":
        return matches
    if occurrence == "first":
        return matches[:1]
    if isinstance(occurrence, int):
        if 0 <= occurrence < len(matches):
            return [matches[occurrence]]
        raise SelectorError(
            f"occurrence={occurrence} out of range: only {len(matches)} match(es) found."
        )
    raise SelectorError(f"Invalid occurrence value: {occurrence!r}")


class SelectorParser:
    @staticmethod
    def parse_word_selector(selector: Union[WordSelector, dict], doc) -> List[Any]:
        if isinstance(selector, dict):
            selector = WordSelector(**selector)
        paragraphs = list(doc.paragraphs)

        if selector.type == "paragraph":
            if selector.index is not None:
                if not 0 <= selector.index < len(paragraphs):
                    raise SelectorError(
                        f"Paragraph index {selector.index} out of range: "
                        f"document has {len(paragraphs)} paragraphs (0..{len(paragraphs)-1})."
                    )
                return [paragraphs[selector.index]]
            if selector.range:
                start, end = selector.range
                if start < 0 or end >= len(paragraphs) or start > end:
                    raise SelectorError(
                        f"Paragraph range ({start}, {end}) invalid: "
                        f"document has {len(paragraphs)} paragraphs."
                    )
                return paragraphs[start : end + 1]
            return paragraphs

        if selector.type == "table":
            tables = doc.tables
            if selector.table_index is None:
                return list(tables)
            if not 0 <= selector.table_index < len(tables):
                raise SelectorError(
                    f"Table index {selector.table_index} out of range: "
                    f"document has {len(tables)} table(s)."
                )
            table = tables[selector.table_index]
            if selector.row_index is None:
                return [table]
            rows = table.rows
            if not 0 <= selector.row_index < len(rows):
                raise SelectorError(
                    f"Row index {selector.row_index} out of range: "
                    f"table {selector.table_index} has {len(rows)} row(s)."
                )
            row = rows[selector.row_index]
            if selector.cell_index is None:
                return [row]
            cells = row.cells
            if not 0 <= selector.cell_index < len(cells):
                raise SelectorError(
                    f"Cell index {selector.cell_index} out of range: "
                    f"row has {len(cells)} cell(s)."
                )
            return [cells[selector.cell_index]]

        if selector.type == "run":
            if selector.index is None:
                raise SelectorError("run selector requires 'index' (paragraph index).")
            if not 0 <= selector.index < len(paragraphs):
                raise SelectorError(
                    f"Paragraph index {selector.index} out of range: "
                    f"document has {len(paragraphs)} paragraphs."
                )
            runs = paragraphs[selector.index].runs
            if selector.run_index is None:
                return list(runs)
            if not 0 <= selector.run_index < len(runs):
                raise SelectorError(
                    f"Run index {selector.run_index} out of range: "
                    f"paragraph {selector.index} has {len(runs)} run(s)."
                )
            return [runs[selector.run_index]]

        if selector.type in ("header", "footer"):
            sections = doc.sections
            if not 0 <= selector.section_index < len(sections):
                raise SelectorError(
                    f"Section index {selector.section_index} out of range: "
                    f"document has {len(sections)} section(s)."
                )
            section = sections[selector.section_index]
            part = section.header if selector.type == "header" else section.footer
            return list(part.paragraphs)

        if selector.type == "text_match":
            if not selector.contains and not selector.regex:
                raise SelectorError(
                    "text_match selector requires 'contains' or 'regex'."
                )
            matches = []
            for para in paragraphs:
                text = para.text
                if selector.contains is not None:
                    needle, hay = selector.contains, text
                    if not selector.case_sensitive:
                        needle, hay = needle.lower(), hay.lower()
                    if needle not in hay:
                        continue
                if selector.regex is not None:
                    flags = 0 if selector.case_sensitive else re.IGNORECASE
                    if not re.search(selector.regex, text, flags):
                        continue
                matches.append(para)
            return _match_occurrence(matches, selector.occurrence)

        if selector.type == "style_match":
            if not selector.style_name:
                raise SelectorError("style_match selector requires 'style_name'.")
            matches = [
                p for p in paragraphs if p.style is not None and p.style.name == selector.style_name
            ]
            return _match_occurrence(matches, selector.occurrence)

        raise SelectorError(f"Unknown Word selector type: {selector.type!r}")

    @staticmethod
    def parse_excel_selector(
        selector: Union[ExcelSelector, dict], workbook
    ) -> Tuple[Any, Optional[tuple]]:
        """Return (worksheet, coords-or-condition)."""
        if isinstance(selector, dict):
            selector = ExcelSelector(**selector)
        if selector.sheet not in workbook.sheetnames:
            raise SelectorError(
                f"Sheet '{selector.sheet}' not found: "
                f"workbook has sheets {workbook.sheetnames}."
            )
        worksheet = workbook[selector.sheet]
        if selector.condition:
            return worksheet, selector.condition
        try:
            coords = selector.parse_range()
        except ValueError as exc:
            raise SelectorError(str(exc)) from exc
        return worksheet, coords
