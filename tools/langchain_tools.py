"""LangChain tools registration and agent integration"""
from typing import List
try:
    from langchain.tools import Tool
except ImportError:
    from langchain_core.tools import Tool

from .base_tools import (
    open_document,
    get_structure,
    apply_changes,
    save_document,
    undo,
    restore_snapshot
)
from .excel_tools import (
    excel_read_cells,
    excel_write_cells,
    excel_edit_formula,
    excel_edit_style,
    excel_insert_rows_cols,
    excel_create_chart,
    excel_conditional_select
)
from .word_tools import (
    word_read_content,
    word_edit_text,
    word_edit_style,
    word_insert_element,
    word_find_replace
)


def get_all_tools() -> List[Tool]:
    """Get all available tools for LangChain agent"""
    return [
        # Base tools
        open_document,
        get_structure,
        apply_changes,
        save_document,
        undo,
        restore_snapshot,
        
        # Excel tools
        excel_read_cells,
        excel_write_cells,
        excel_edit_formula,
        excel_edit_style,
        excel_insert_rows_cols,
        excel_create_chart,
        excel_conditional_select,
        
        # Word tools
        word_read_content,
        word_edit_text,
        word_edit_style,
        word_insert_element,
        word_find_replace,
    ]


def get_excel_tools() -> List[Tool]:
    """Get only Excel tools"""
    return [
        excel_read_cells,
        excel_write_cells,
        excel_edit_formula,
        excel_edit_style,
        excel_insert_rows_cols,
        excel_create_chart,
        excel_conditional_select,
    ]


def get_word_tools() -> List[Tool]:
    """Get only Word tools"""
    return [
        word_read_content,
        word_edit_text,
        word_edit_style,
        word_insert_element,
        word_find_replace,
    ]


def get_base_tools() -> List[Tool]:
    """Get only base document management tools"""
    return [
        open_document,
        get_structure,
        apply_changes,
        save_document,
        undo,
        restore_snapshot,
    ]
