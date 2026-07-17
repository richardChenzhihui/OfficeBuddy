"""Word editing tools. All mutations apply directly to the working copy and are
snapshotted automatically; there is no separate preview/apply phase."""
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field

from ..adapters.word_adapter import WordAdapter
from ..core.selector_parser import SelectorParser
from ..schemas.operations import StyleParams, TextOperation
from .registry import REGISTRY, ToolContext

_SELECTOR_DOC = (
    "Word selector dict. Types: "
    "{'type':'paragraph','index':N} or {'type':'paragraph','range':[start,end]}; "
    "{'type':'text_match','contains':'...','case_sensitive':false,'occurrence':'first'|'all'|N}; "
    "{'type':'style_match','style_name':'Heading 1'}; "
    "{'type':'table','table_index':N,'row_index':N,'cell_index':N}; "
    "{'type':'run','index':paragraph_idx,'run_index':N}; "
    "{'type':'header'|'footer','section_index':0}. "
    "Prefer text_match over raw indices when targeting content by meaning."
)


def _word_session(ctx: ToolContext, doc_id: str):
    session = ctx.sessions.get(doc_id)
    if session.doc_type != "word":
        raise ValueError(
            f"doc_id '{doc_id}' is an Excel document; use excel_* tools instead."
        )
    return session


def _resolve(session, selector: dict) -> List[Any]:
    elements = SelectorParser.parse_word_selector(selector, session.doc)
    if not elements:
        raise ValueError(
            f"Selector matched no elements: {selector}. "
            f"Document has {len(session.doc.paragraphs)} paragraphs and "
            f"{len(session.doc.tables)} table(s). Call get_structure to inspect, "
            "or loosen the match (e.g. case_sensitive=false)."
        )
    return elements


class WordReadInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    selector: Dict[str, Any] = Field(..., description=_SELECTOR_DOC)


class WordEditTextInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    selector: Dict[str, Any] = Field(..., description=_SELECTOR_DOC)
    operation: TextOperation = Field(
        ...,
        description=(
            "replace: set element text; insert: prepend at element start; "
            "delete: clear text; append: add at element end (keeps formatting)"
        ),
    )
    text: Optional[str] = Field(None, description="Text for replace/insert/append")


class WordEditStyleInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    selector: Dict[str, Any] = Field(..., description=_SELECTOR_DOC)
    style_params: StyleParams = Field(
        ...,
        description=(
            "Formatting to apply, e.g. {'font_name':'Times New Roman',"
            "'font_size':12.0,'bold':true,'alignment':'center'}"
        ),
    )


class WordInsertElementInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    element_type: str = Field(
        ..., description="Element to insert: 'paragraph', 'table', or 'page_break'"
    )
    content: Optional[Any] = Field(
        None,
        description=(
            "paragraph: string; table: list of row lists e.g. [['a','b'],['c','d']]"
        ),
    )
    position: Optional[int] = Field(
        None,
        description=(
            "Paragraph index to insert BEFORE (applies to both paragraphs and "
            "tables). Omit to append at document end."
        ),
    )


class WordFindReplaceInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    find: str = Field(..., description="Exact text to find")
    replace: str = Field(..., description="Replacement text")
    scope: Optional[Dict[str, Any]] = Field(
        None, description=f"Optional scope selector. {_SELECTOR_DOC}"
    )


@REGISTRY.register(
    "word_read_content",
    "Read full text/content of the selected Word elements (untruncated).",
    WordReadInput,
)
def word_read_content(ctx: ToolContext, p: WordReadInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    elements = SelectorParser.parse_word_selector(p.selector, session.doc)
    if not elements:
        return {
            "content": [],
            "matched_count": 0,
            "note": f"Selector matched no elements: {p.selector}",
        }
    return {
        "content": WordAdapter.read_content(elements),
        "matched_count": len(elements),
    }


@REGISTRY.register(
    "word_edit_text",
    "Edit text of selected Word elements. Applies immediately to the working copy "
    "(original file untouched); each call is snapshotted and undoable.",
    WordEditTextInput,
    mutates=True,
)
def word_edit_text(ctx: ToolContext, p: WordEditTextInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    if p.operation in (TextOperation.REPLACE, TextOperation.INSERT, TextOperation.APPEND):
        if not p.text:
            raise ValueError(f"operation '{p.operation.value}' requires 'text'.")
    elements = _resolve(session, p.selector)
    result = WordAdapter.edit_text(elements, p.operation, p.text)
    return {"matched_count": len(elements), **result}


@REGISTRY.register(
    "word_edit_style",
    "Apply formatting (font, size, bold, color, alignment, spacing) to selected "
    "Word elements. Applies immediately; snapshotted and undoable.",
    WordEditStyleInput,
    mutates=True,
)
def word_edit_style(ctx: ToolContext, p: WordEditStyleInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    elements = _resolve(session, p.selector)
    result = WordAdapter.apply_style(elements, p.style_params)
    return {"matched_count": len(elements), **result}


@REGISTRY.register(
    "word_insert_element",
    "Insert a new paragraph, table, or page break. Applies immediately; "
    "snapshotted and undoable.",
    WordInsertElementInput,
    mutates=True,
)
def word_insert_element(ctx: ToolContext, p: WordInsertElementInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    return WordAdapter.insert_element(
        session.doc, p.position, p.element_type, p.content
    )


class WordDeleteElementInput(BaseModel):
    doc_id: str = Field(..., description="Document id")
    selector: Dict[str, Any] = Field(
        ...,
        description=(
            "Selector for the element(s) to delete (paragraph / table / "
            f"text_match / style_match). {_SELECTOR_DOC}"
        ),
    )


@REGISTRY.register(
    "word_delete_element",
    "Delete the selected paragraph(s), table(s), or table row(s) from the "
    "document entirely. Use this (or undo) to remove a duplicated or wrongly "
    "inserted element — never insert again hoping to fix it.",
    WordDeleteElementInput,
    mutates=True,
)
def word_delete_element(ctx: ToolContext, p: WordDeleteElementInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    elements = _resolve(session, p.selector)
    result = WordAdapter.delete_elements(elements)
    return {"matched_count": len(elements), **result}


@REGISTRY.register(
    "word_find_replace",
    "Find and replace exact text across the document or a scoped selection. "
    "Preserves run formatting when matches don't span formatting boundaries.",
    WordFindReplaceInput,
    mutates=True,
)
def word_find_replace(ctx: ToolContext, p: WordFindReplaceInput) -> dict:
    session = _word_session(ctx, p.doc_id)
    scope_elements = None
    if p.scope:
        scope_elements = _resolve(session, p.scope)
    result = WordAdapter.find_replace(session.doc, p.find, p.replace, scope_elements)
    if result["count"] == 0:
        return {
            "success": False,
            **result,
            "error": (
                f"'{p.find}' not found"
                + (" in the given scope" if p.scope else " anywhere in the document")
                + ". Check exact spelling/whitespace via get_structure or "
                "word_read_content."
            ),
        }
    return result
