"""Base tools for document management"""
import uuid
from typing import Dict, Any, Optional
try:
    from langchain.tools import tool
except ImportError:
    from langchain_core.tools import tool
from pydantic import BaseModel, Field

from ..core.document_manager import DocumentManager


class OpenDocumentInput(BaseModel):
    """Input for open_document tool"""
    file_path: str = Field(..., description="Path to the document file")


class GetStructureInput(BaseModel):
    """Input for get_structure tool"""
    doc_id: str = Field(..., description="Document ID")
    depth: int = Field(2, description="Structure depth to retrieve")


class ApplyChangesInput(BaseModel):
    """Input for apply_changes tool"""
    doc_id: str = Field(..., description="Document ID")
    preview_id: str = Field(..., description="Preview ID from preview operation")


class SaveDocumentInput(BaseModel):
    """Input for save_document tool"""
    doc_id: str = Field(..., description="Document ID")
    path: Optional[str] = Field(None, description="Save path (optional, uses original path if not provided)")


class UndoInput(BaseModel):
    """Input for undo tool"""
    doc_id: str = Field(..., description="Document ID")


class RestoreSnapshotInput(BaseModel):
    """Input for restore_snapshot tool"""
    doc_id: str = Field(..., description="Document ID")
    snapshot_id: str = Field(..., description="Snapshot ID to restore")


# Global document manager instance
_doc_manager = DocumentManager()


@tool(args_schema=OpenDocumentInput)
def open_document(file_path: str) -> Dict[str, Any]:
    """Open a Word or Excel document and return document ID and summary.
    
    Args:
        file_path: Path to the document file (.docx, .xlsx, .xlsm)
    
    Returns:
        Dictionary with doc_id and summary information
    """
    result = _doc_manager.open_document(file_path)
    return result


@tool(args_schema=GetStructureInput)
def get_structure(doc_id: str, depth: int = 2) -> Dict[str, Any]:
    """Get the structure of a document.
    
    Args:
        doc_id: Document ID from open_document
        depth: Depth of structure to retrieve (default: 2)
    
    Returns:
        Document structure tree
    """
    structure = _doc_manager.get_structure(doc_id, depth)
    return structure


@tool(args_schema=ApplyChangesInput)
def apply_changes(doc_id: str, preview_id: str) -> Dict[str, Any]:
    """Apply previously previewed changes to the document.
    
    IMPORTANT: Both doc_id and preview_id must be exact UUID strings from previous tool responses.
    Do NOT format them as "doc_id: value" - pass the UUID string directly.
    
    Args:
        doc_id: Document ID (UUID string from open_document response, e.g., "22a6941f-14dd-4c26-9293-f58d3ab8de2b")
        preview_id: Preview ID (UUID string from edit tool response, e.g., "3ca6f526-3e65-456c-86a9-e61d9fbb45d9")
    
    Returns:
        Dictionary with success status and snapshot_id
    
    Note: If you have multiple previews, call this function separately for each preview_id.
    """
    result = _doc_manager.apply_changes(doc_id, preview_id)
    return result


@tool(args_schema=SaveDocumentInput)
def save_document(doc_id: str, path: Optional[str] = None) -> Dict[str, Any]:
    """Save the document to disk.
    
    Args:
        doc_id: Document ID
        path: Optional save path (uses original path if not provided)
    
    Returns:
        Dictionary with success status
    """
    success = _doc_manager.save_document(doc_id, path)
    return {"success": success}


@tool(args_schema=UndoInput)
def undo(doc_id: str) -> Dict[str, Any]:
    """Undo the last change to the document.
    
    Args:
        doc_id: Document ID
    
    Returns:
        Dictionary with success status
    """
    success = _doc_manager.undo(doc_id)
    return {"success": success}


@tool(args_schema=RestoreSnapshotInput)
def restore_snapshot(doc_id: str, snapshot_id: str) -> Dict[str, Any]:
    """Restore document from a specific snapshot.
    
    Args:
        doc_id: Document ID
        snapshot_id: Snapshot ID to restore
    
    Returns:
        Dictionary with success status
    """
    success = _doc_manager.restore_snapshot(doc_id, snapshot_id)
    return {"success": success}


def get_document_manager() -> DocumentManager:
    """Get the global document manager instance"""
    return _doc_manager
