"""Word adapter wrapping python-docx. Stateless: mutates the elements passed in."""
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

from ..schemas.operations import StyleParams, TextOperation

_ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _copy_run_format(src: Any, dst: Any) -> None:
    dst.font.name = src.font.name
    if src.font.size:
        dst.font.size = src.font.size
    dst.font.bold = src.font.bold
    dst.font.italic = src.font.italic
    dst.font.underline = src.font.underline
    if src.font.color and src.font.color.rgb:
        dst.font.color.rgb = src.font.color.rgb


def _parse_color(value: str) -> RGBColor:
    raw = value[1:] if value.startswith("#") else value
    try:
        return RGBColor.from_string(raw)
    except Exception as exc:
        raise ValueError(
            f"Invalid color '{value}': expected hex like '#FF0000' or 'FF0000'."
        ) from exc


class WordAdapter:
    """Stateless operations on python-docx elements."""

    @staticmethod
    def read_content(elements: List[Any]) -> List[Dict[str, Any]]:
        result = []
        for element in elements:
            if hasattr(element, "rows"):  # Table
                result.append(
                    {
                        "type": "table",
                        "data": [[cell.text for cell in row.cells] for row in element.rows],
                    }
                )
            elif hasattr(element, "text"):  # Paragraph or Run
                result.append(
                    {
                        "type": "text",
                        "text": element.text,
                        "style": element.style.name if hasattr(element, "style") else None,
                    }
                )
        return result

    @staticmethod
    def edit_text(
        elements: List[Any], operation: TextOperation, text: Optional[str] = None
    ) -> Dict[str, Any]:
        affected: List[str] = []
        for element in elements:
            if not hasattr(element, "text"):
                continue
            old_text = element.text

            if operation == TextOperation.REPLACE:
                element.text = text or ""
                affected.append(f"replaced '{old_text[:50]}' -> '{(text or '')[:50]}'")

            elif operation == TextOperation.INSERT:
                if not text:
                    continue
                if hasattr(element, "runs") and element.runs:
                    # Prepend a run, preserving the paragraph's existing runs/formatting.
                    first = element.runs[0]
                    new_run = element.add_run(text)
                    _copy_run_format(first, new_run)
                    first._element.addprevious(new_run._element)
                elif hasattr(element, "runs"):
                    element.add_run(text)
                else:  # Run
                    element.text = text + element.text
                affected.append(f"inserted '{text[:50]}' at start")

            elif operation == TextOperation.DELETE:
                element.text = ""
                affected.append(f"deleted '{old_text[:50]}'")

            elif operation == TextOperation.APPEND:
                if not text:
                    continue
                if hasattr(element, "runs"):
                    if element.runs:
                        last_run = element.runs[-1]
                        new_run = element.add_run(text)
                        _copy_run_format(last_run, new_run)
                    else:
                        element.add_run(text)
                else:  # Run
                    element.text = element.text + text
                affected.append(f"appended '{text[:50]}'")

        return {"affected": affected, "operation": operation.value}

    @staticmethod
    def apply_style(elements: List[Any], style: StyleParams) -> Dict[str, Any]:
        affected: List[str] = []
        for element in elements:
            if hasattr(element, "runs"):  # Paragraph: style runs AND paragraph format
                for run in element.runs:
                    WordAdapter._apply_run_style(run, style)
                WordAdapter._apply_paragraph_format(element, style)
                affected.append(f"paragraph styled ({len(element.runs)} runs)")
            elif hasattr(element, "font"):  # Run
                WordAdapter._apply_run_style(element, style)
                affected.append("run styled")
        return {"affected": affected}

    @staticmethod
    def _apply_paragraph_format(paragraph: Any, style: StyleParams) -> None:
        fmt = paragraph.paragraph_format
        if style.paragraph_spacing is not None:
            fmt.space_before = Pt(style.paragraph_spacing)
        if style.line_spacing is not None:
            fmt.line_spacing = style.line_spacing
        if style.alignment:
            if style.alignment not in _ALIGNMENT_MAP:
                raise ValueError(
                    f"Invalid alignment '{style.alignment}': "
                    f"expected one of {sorted(_ALIGNMENT_MAP)}."
                )
            fmt.alignment = _ALIGNMENT_MAP[style.alignment]

    @staticmethod
    def _apply_run_style(run: Any, style: StyleParams) -> None:
        font = run.font
        if style.font_name:
            font.name = style.font_name
        if style.font_size:
            font.size = Pt(style.font_size)
        if style.bold is not None:
            font.bold = style.bold
        if style.italic is not None:
            font.italic = style.italic
        if style.underline is not None:
            font.underline = style.underline
        if style.color:
            font.color.rgb = _parse_color(style.color)

    @staticmethod
    def insert_element(
        doc: Document, position: Optional[int], element_type: str, content: Any
    ) -> Dict[str, Any]:
        if element_type == "table":
            if not isinstance(content, list) or not all(
                isinstance(row, list) for row in content
            ):
                raise ValueError("Table content must be a list of row lists.")
            cols = len(content[0]) if content else 0
            table = doc.add_table(rows=len(content), cols=cols)
            for i, row_data in enumerate(content):
                for j, cell_data in enumerate(row_data):
                    table.rows[i].cells[j].text = str(cell_data)
            if position is not None:
                if not 0 <= position < len(doc.paragraphs):
                    raise ValueError(
                        f"Insert position {position} out of range: "
                        f"document has {len(doc.paragraphs)} paragraphs."
                    )
                doc.paragraphs[position]._element.addnext(table._element)
            return {"element_type": "table", "rows": len(content), "cols": cols}

        if element_type == "paragraph":
            if position is not None:
                if not 0 <= position < len(doc.paragraphs):
                    raise ValueError(
                        f"Insert position {position} out of range: "
                        f"document has {len(doc.paragraphs)} paragraphs."
                    )
                doc.paragraphs[position].insert_paragraph_before(str(content))
            else:
                doc.add_paragraph(str(content))
            return {"element_type": "paragraph", "text": str(content)[:50]}

        if element_type == "page_break":
            from docx.enum.text import WD_BREAK

            if position is not None and 0 <= position < len(doc.paragraphs):
                para = doc.paragraphs[position]
            else:
                para = doc.add_paragraph()
            para.add_run().add_break(WD_BREAK.PAGE)
            return {"element_type": "page_break"}

        raise ValueError(
            f"Unsupported element type '{element_type}': "
            "expected table, paragraph, or page_break."
        )

    @staticmethod
    def find_replace(
        doc: Document, find: str, replace: str, scope: Optional[List[Any]] = None
    ) -> Dict[str, Any]:
        if not find:
            raise ValueError("find string must be non-empty.")
        count = 0
        affected: List[str] = []
        formatting_lost_in: List[str] = []
        elements = scope if scope else doc.paragraphs

        for element in elements:
            if not hasattr(element, "text"):
                continue
            occurrences = element.text.count(find)
            if not occurrences:
                continue
            # Replace inside individual runs first to preserve formatting.
            if hasattr(element, "runs"):
                for run in element.runs:
                    if find in run.text:
                        run.text = run.text.replace(find, replace)
                if find in element.text:
                    # Occurrences spanning run boundaries: paragraph-level fallback,
                    # which flattens the paragraph's run formatting.
                    element.text = element.text.replace(find, replace)
                    formatting_lost_in.append(element.text[:50])
            else:
                element.text = element.text.replace(find, replace)
            count += occurrences
            affected.append(f"'{element.text[:50]}'")

        result: Dict[str, Any] = {
            "count": count,
            "affected": affected,
            "find": find,
            "replace": replace,
        }
        if formatting_lost_in:
            result["warning"] = (
                "Some matches spanned formatting boundaries; run formatting was "
                f"flattened in {len(formatting_lost_in)} paragraph(s)."
            )
        return result
