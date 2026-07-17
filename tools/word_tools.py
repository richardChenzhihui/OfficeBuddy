"""Word editing tools"""
import uuid
from typing import Dict, Any, Optional, List
try:
    from langchain.tools import tool
except ImportError:
    from langchain_core.tools import tool
from pydantic import BaseModel, Field

from .base_tools import get_document_manager
from ..core.selector_parser import SelectorParser
from ..adapters.word_adapter import WordAdapter
from ..schemas.selector import WordSelector
from ..schemas.operations import StyleParams, TextOperation


class WordReadContentInput(BaseModel):
    """Input for word_read_content tool"""
    doc_id: str = Field(..., description="Document ID")
    selector: Dict[str, Any] = Field(..., description="Word selector dictionary")


class WordEditTextInput(BaseModel):
    """Input for word_edit_text tool"""
    doc_id: str = Field(..., description="Document ID")
    selector: Dict[str, Any] = Field(..., description="Word selector dictionary")
    operation: str = Field(..., description="Operation: replace, insert, delete, append")
    text: Optional[str] = Field(None, description="Text for operation")


class WordEditStyleInput(BaseModel):
    """Input for word_edit_style tool"""
    doc_id: str = Field(..., description="Document ID")
    selector: Dict[str, Any] = Field(..., description="Word selector dictionary")
    style_params: Dict[str, Any] = Field(..., description="Style parameters")


class WordInsertElementInput(BaseModel):
    """Input for word_insert_element tool"""
    doc_id: str = Field(..., description="Document ID")
    position: Optional[int] = Field(None, description="Insert position (paragraph index)")
    element_type: str = Field(..., description="Element type: table, paragraph, page_break")
    content: Any = Field(..., description="Content for the element")


class WordFindReplaceInput(BaseModel):
    """Input for word_find_replace tool"""
    doc_id: str = Field(..., description="Document ID")
    find: str = Field(..., description="Text to find")
    replace: str = Field(..., description="Replacement text")
    scope: Optional[Dict[str, Any]] = Field(None, description="Optional scope selector")


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


@tool(args_schema=WordReadContentInput)
def word_read_content(doc_id: str, selector: Dict[str, Any]) -> Dict[str, Any]:
    """Read content from Word document.
    
    Args:
        doc_id: Document ID
        selector: Word selector dictionary (type, index, range, etc.)
    
    Returns:
        Dictionary with content list
    """
    doc = _doc_manager.get_document(doc_id)
    word_selector = WordSelector(**selector)
    elements = _selector_parser.parse_word_selector(word_selector, doc)
    
    result = WordAdapter.read_content(elements)
    return {
        "content": result,
        "count": len(result)
    }


@tool(args_schema=WordEditTextInput)
def word_edit_text(doc_id: str, selector: Dict[str, Any], operation: str, 
                   text: Optional[str] = None) -> Dict[str, Any]:
    """Edit text in Word document (preview mode).
    
    Args:
        doc_id: Document ID
        selector: Word selector dictionary
        operation: Operation type (replace, insert, delete, append)
        text: Text for operation (required for replace, insert, append)
    
    Returns:
        Dictionary with preview_id and preview information
    """
    doc = _doc_manager.get_document(doc_id)
    word_selector = WordSelector(**selector)
    elements = _selector_parser.parse_word_selector(word_selector, doc)
    
    if not elements:
        return {"error": "No elements found matching selector"}
    
    try:
        operation_enum = TextOperation(operation)
    except ValueError:
        return {"error": f"Invalid operation: {operation}"}
    
    result = WordAdapter.edit_text(elements, operation_enum, text)
    
    preview_id = _create_preview(doc_id, "word_edit_text", result, {
        "selector": selector,
        "operation": operation,
        "text": text
    })
    
    return {
        "preview_id": preview_id,
        "preview": result,
        "diff": f"Will {operation} text in {len(result.get('affected', []))} element(s)"
    }


@tool(args_schema=WordEditStyleInput)
def word_edit_style(doc_id: str, selector: Dict[str, Any], 
                    style_params: Dict[str, Any]) -> Dict[str, Any]:
    """Edit style in Word document (preview mode).
    
    Available style parameters in style_params:
    - font_name: str - Font name (e.g., "Times New Roman", "Arial", "Calibri")
    - font_size: float - Font size in points (e.g., 12.0, 14.0, 16.0)
    - bold: bool - Bold text (True/False)
    - italic: bool - Italic text (True/False)
    - underline: bool - Underline text (True/False)
    - color: str - Text color in hex format (e.g., "#FF0000" for red) or RGB
    - bg_color: str - Background color in hex format
    - alignment: str - Paragraph alignment: "left", "center", "right", "justify"
    - paragraph_spacing: float - Space before paragraph in points
    - line_spacing: float - Line spacing multiplier
    
    Example style_params:
    {
        "font_name": "Times New Roman",
        "font_size": 12.0,
        "bold": True,
        "italic": False,
        "underline": True,
        "color": "#000000"
    }
    
    Args:
        doc_id: Document ID
        selector: Word selector dictionary (e.g., {"type": "paragraph", "index": 0})
        style_params: Style parameters dictionary with any of the above parameters
    
    Returns:
        Dictionary with preview_id and preview information
    """
    doc = _doc_manager.get_document(doc_id)
    word_selector = WordSelector(**selector)
    elements = _selector_parser.parse_word_selector(word_selector, doc)
    
    if not elements:
        return {"error": "No elements found matching selector"}
    
    style = StyleParams(**style_params)
    result = WordAdapter.apply_style(elements, style)
    
    preview_id = _create_preview(doc_id, "word_edit_style", result, {
        "selector": selector,
        "style_params": style_params
    })
    
    return {
        "preview_id": preview_id,
        "preview": result,
        "diff": f"Will apply style to {len(result.get('affected', []))} element(s)"
    }


@tool(args_schema=WordInsertElementInput)
def word_insert_element(doc_id: str, element_type: str, content: Any, 
                       position: Optional[int] = None) -> Dict[str, Any]:
    """Insert element into Word document (preview mode).
    
    Args:
        doc_id: Document ID
        element_type: Element type (table, paragraph, page_break)
        content: Content for the element
        position: Optional insert position (paragraph index)
    
    Returns:
        Dictionary with preview_id and preview information
    """
    # Create preview without actually modifying the document
    # The actual insertion will happen in apply_changes
    if element_type == "paragraph":
        preview_text = str(content)[:50] if content else ""
        result = {
            "element_type": "paragraph",
            "text": preview_text
        }
    elif element_type == "table":
        if isinstance(content, list):
            result = {
                "element_type": "table",
                "rows": len(content),
                "cols": len(content[0]) if content else 0
            }
        else:
            return {"error": "Table content must be a list of lists"}
    elif element_type == "page_break":
        result = {
            "element_type": "page_break"
        }
    else:
        return {"error": f"Unsupported element type: {element_type}"}
    
    preview_id = _create_preview(doc_id, "word_insert_element", result, {
        "position": position,
        "element_type": element_type,
        "content": content
    })
    
    return {
        "preview_id": preview_id,
        "preview": result,
        "diff": f"Will insert {element_type} at position {position if position is not None else 'end'}"
    }


@tool(args_schema=WordFindReplaceInput)
def word_find_replace(doc_id: str, find: str, replace: str, 
                     scope: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Find and replace text in Word document (preview mode).
    
    Args:
        doc_id: Document ID
        find: Text to find
        replace: Replacement text
        scope: Optional scope selector to limit search
    
    Returns:
        Dictionary with preview_id, count, and preview information
    """
    doc = _doc_manager.get_document(doc_id)
    
    elements = None
    if scope:
        word_selector = WordSelector(**scope)
        elements = _selector_parser.parse_word_selector(word_selector, doc)
    
    result = WordAdapter.find_replace(doc, find, replace, elements)
    
    preview_id = _create_preview(doc_id, "word_find_replace", result, {
        "find": find,
        "replace": replace,
        "scope": scope
    })
    
    return {
        "preview_id": preview_id,
        "count": result.get("count", 0),
        "preview": result,
        "diff": f"Will replace '{find}' with '{replace}' {result.get('count', 0)} time(s)"
    }
