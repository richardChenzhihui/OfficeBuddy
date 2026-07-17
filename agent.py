"""LangChain agent setup for Office document editing"""
import os
from typing import Optional
try:
    from langchain.agents import AgentExecutor, create_openai_tools_agent
    from langchain.prompts import ChatPromptTemplate, MessagesPlaceholder
except ImportError:
    from langchain.agents import AgentExecutor
    from langchain.agents import create_openai_tools_agent
    from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_openai import ChatOpenAI
from langchain_core.messages import SystemMessage

from .tools.langchain_tools import get_all_tools


def create_office_agent(
    model_name: str = "openai/gpt-4",
    temperature: float = 0.0,
    api_key: Optional[str] = None,
    base_url: Optional[str] = None,
    use_openrouter: bool = True
) -> AgentExecutor:
    """Create a LangChain agent for Office document editing
    
    Args:
        model_name: Model name (default: openai/gpt-4 for OpenRouter)
        temperature: Temperature for LLM (default: 0.0)
        api_key: API key (optional, uses OPENROUTER_API_KEY env var if use_openrouter=True, 
                 otherwise uses OPENAI_API_KEY)
        base_url: Base URL for API (optional, defaults to OpenRouter if use_openrouter=True)
        use_openrouter: Whether to use OpenRouter (default: True)
    
    Returns:
        AgentExecutor instance
    """
    # Initialize LLM with OpenRouter by default
    if use_openrouter:
        # Default to OpenRouter configuration
        if base_url is None:
            base_url = "https://openrouter.ai/api/v1"
        if api_key is None:
            api_key = os.getenv("OPENROUTER_API_KEY")
            if not api_key:
                raise ValueError(
                    "OPENROUTER_API_KEY environment variable not set. "
                    "Please set it or provide api_key parameter."
                )
    else:
        # Use OpenAI or custom endpoint
        if api_key is None:
            api_key = os.getenv("OPENAI_API_KEY")
    
    llm_kwargs = {
        "model": model_name,
        "temperature": temperature,
    }
    if api_key:
        llm_kwargs["api_key"] = api_key
    if base_url:
        llm_kwargs["base_url"] = base_url
    
    llm = ChatOpenAI(**llm_kwargs)
    
    # Get tools
    tools = get_all_tools()
    
    # Create prompt
    system_message = SystemMessage(content="""You are an expert assistant for editing Word and Excel documents.

Key principles:
1. Always use structured selectors (not natural language) to locate elements
2. Preview changes before applying them - use preview_id from preview operations
3. Use get_structure to understand document layout before editing
4. Apply changes using apply_changes with preview_id
5. Save documents after making changes

Workflow:
1. open_document(file_path) - Returns {"doc_id": "uuid-string", "summary": {...}}
2. get_structure(doc_id) - Understand document structure
3. Use appropriate read/edit tools with structured selectors
   - Each edit tool returns {"preview_id": "uuid-string", "preview": {...}, ...}
4. IMPORTANT: When you have multiple preview operations, apply each one separately:
   - apply_changes(doc_id="the-exact-doc-id-string", preview_id="the-exact-preview-id-string")
   - Call apply_changes once for each preview_id you created
5. save_document(doc_id) - Save the document

CRITICAL: When calling apply_changes:
- Extract the exact doc_id string from open_document response (e.g., "22a6941f-14dd-4c26-9293-f58d3ab8de2b")
- Extract the exact preview_id string from each edit tool response (e.g., "3ca6f526-3e65-456c-86a9-e61d9fbb45d9")
- Pass these as direct string values, NOT as formatted text like "doc_id: value"
- If you have multiple previews, call apply_changes separately for each preview_id

For Word documents:
- Use word_read_content with selector: {"type": "paragraph", "index": 0}
- Use word_edit_text with operation: replace/insert/delete/append
- Use word_edit_style to modify formatting (font_name, font_size, bold, italic, underline, color, alignment, etc.)
  Example: {"font_name": "Times New Roman", "font_size": 12.0, "bold": True, "underline": True}

For Excel documents:
- Use excel_read_cells with sheet and range (e.g., "A1:B10")
- Use excel_write_cells with 2D array of values
- Use excel_edit_formula to set formulas
- Use excel_edit_style to modify formatting (font_name, font_size, bold, italic, underline, color, bg_color, alignment, etc.)
  Example: {"font_name": "Arial", "font_size": 14.0, "bold": True, "bg_color": "#FFFF00"}

Always be careful to preserve document structure and formatting.""")
    
    prompt = ChatPromptTemplate.from_messages([
        system_message,
        MessagesPlaceholder(variable_name="chat_history", optional=True),
        ("human", "{input}"),
        MessagesPlaceholder(variable_name="agent_scratchpad"),
    ])
    
    # Create agent
    agent = create_openai_tools_agent(llm, tools, prompt)
    
    # Create executor
    executor = AgentExecutor(
        agent=agent,
        tools=tools,
        verbose=True,
        handle_parsing_errors=True,
        max_iterations=15
    )
    
    return executor
