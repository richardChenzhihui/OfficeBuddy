"""Selector parser for converting selectors to document element references"""
from typing import Any, Optional, Union
from ..schemas.selector import WordSelector, ExcelSelector, Selector


class SelectorParser:
    """Parses selectors and converts them to document element references"""
    
    @staticmethod
    def parse_word_selector(selector: Union[WordSelector, dict], doc) -> Any:
        """Parse Word selector and return document elements"""
        if isinstance(selector, dict):
            selector = WordSelector(**selector)
        
        if selector.type == "paragraph":
            paragraphs = list(doc.paragraphs)
            if selector.index is not None:
                if 0 <= selector.index < len(paragraphs):
                    return [paragraphs[selector.index]]
            elif selector.range:
                start, end = selector.range
                return paragraphs[start:end+1]
            else:
                return paragraphs
        
        elif selector.type == "table":
            if selector.table_index is not None:
                tables = doc.tables
                if 0 <= selector.table_index < len(tables):
                    table = tables[selector.table_index]
                    if selector.row_index is not None:
                        rows = table.rows
                        if 0 <= selector.row_index < len(rows):
                            row = rows[selector.row_index]
                            if selector.cell_index is not None:
                                cells = row.cells
                                if 0 <= selector.cell_index < len(cells):
                                    return [cells[selector.cell_index]]
                                return []
                            return [row]
                        return []
                    return [table]
            return doc.tables
        
        elif selector.type == "run":
            if selector.index is not None:
                paragraphs = list(doc.paragraphs)
                if 0 <= selector.index < len(paragraphs):
                    para = paragraphs[selector.index]
                    runs = para.runs
                    if selector.run_index is not None:
                        if 0 <= selector.run_index < len(runs):
                            return [runs[selector.run_index]]
                    return runs
            return []
        
        elif selector.type in ["header", "footer"]:
            section = doc.sections[0]
            if selector.type == "header":
                return [section.header]
            else:
                return [section.footer]
        
        return []
    
    @staticmethod
    def parse_excel_selector(selector: Union[ExcelSelector, dict], workbook) -> tuple[Any, Optional[tuple]]:
        """Parse Excel selector and return worksheet and range coordinates"""
        if isinstance(selector, dict):
            selector = ExcelSelector(**selector)
        
        sheet_name = selector.sheet
        if sheet_name not in workbook.sheetnames:
            return None, None
        
        worksheet = workbook[sheet_name]
        
        if selector.condition:
            # Conditional selection - return all matching cells
            return worksheet, selector.condition
        
        # Parse range
        sheet_name, coords = selector.parse_range()
        return worksheet, coords
    
    @staticmethod
    def parse_selector(selector: Union[Selector, dict], doc) -> Any:
        """Parse unified selector"""
        if isinstance(selector, dict):
            if "doc_type" not in selector:
                raise ValueError("Selector must have doc_type field")
            selector = Selector(**selector)
        
        if selector.doc_type == "word":
            if not selector.word:
                raise ValueError("Word selector requires word field")
            return SelectorParser.parse_word_selector(selector.word, doc)
        elif selector.doc_type == "excel":
            if not selector.excel:
                raise ValueError("Excel selector requires excel field")
            return SelectorParser.parse_excel_selector(selector.excel, doc)
        else:
            raise ValueError(f"Unknown doc_type: {selector.doc_type}")
