"""Tool registry: pydantic input models -> Anthropic tool schemas + dispatch.

Every tool handler receives (ctx, validated_params_model) and returns a dict.
Dispatch guarantees a uniform result envelope:
  {"success": bool, ..., "error": str|None}
Mutating tools automatically snapshot the document after a successful call,
so undo granularity matches individual tool calls.
"""
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Type

from pydantic import BaseModel, ValidationError

from ..core.selector_parser import SelectorError
from ..core.session import SessionManager
from ..core.snapshot_manager import SnapshotManager


class ToolContext:
    """Shared state handed to every tool handler."""

    def __init__(self) -> None:
        self.sessions = SessionManager()
        self.snapshots: Dict[str, SnapshotManager] = {}
        self.extra: Dict[str, Any] = {}  # renderer, ui hooks, config

    def snapshot_manager(self, doc_id: str) -> SnapshotManager:
        if doc_id not in self.snapshots:
            self.snapshots[doc_id] = SnapshotManager(self.sessions.get(doc_id))
        return self.snapshots[doc_id]


@dataclass
class ToolDef:
    name: str
    description: str
    input_model: Type[BaseModel]
    handler: Callable[[ToolContext, BaseModel], Dict[str, Any]]
    mutates: bool = False


@dataclass
class ToolRegistry:
    tools: Dict[str, ToolDef] = field(default_factory=dict)

    def register(
        self,
        name: str,
        description: str,
        input_model: Type[BaseModel],
        mutates: bool = False,
    ):
        def decorator(func):
            self.tools[name] = ToolDef(name, description, input_model, func, mutates)
            return func

        return decorator

    def to_anthropic_tools(self, names: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        selected = (
            self.tools.values()
            if names is None
            else [self.tools[n] for n in names]
        )
        return [
            {
                "name": t.name,
                "description": t.description,
                "input_schema": t.input_model.model_json_schema(),
            }
            for t in selected
        ]

    def dispatch(self, ctx: ToolContext, name: str, tool_input: Dict[str, Any]) -> Dict[str, Any]:
        if name not in self.tools:
            return {
                "success": False,
                "error": f"Unknown tool '{name}': available tools are {sorted(self.tools)}.",
            }
        tool = self.tools[name]
        try:
            params = tool.input_model(**tool_input)
        except ValidationError as exc:
            return {
                "success": False,
                "error": f"Invalid arguments for {name}: {exc.errors(include_url=False)}",
            }
        try:
            result = tool.handler(ctx, params)
        except (SelectorError, ValueError, KeyError, PermissionError, FileNotFoundError) as exc:
            return {"success": False, "error": str(exc), "error_type": type(exc).__name__}
        except Exception as exc:  # noqa: BLE001 — the envelope contract beats purity:
            # every tool call must return {success, error}, never raise.
            return {
                "success": False,
                "error": f"Unexpected {type(exc).__name__} in {name}: {exc}",
                "error_type": type(exc).__name__,
            }

        result.setdefault("success", True)
        doc_id = getattr(params, "doc_id", None)
        if doc_id and result["success"]:
            session = ctx.sessions.sessions.get(doc_id)
            if session is not None:
                session.audit(name, {"input": tool_input, "result_keys": sorted(result)})
                if tool.mutates:
                    session.dirty = True
                    snap_id = ctx.snapshot_manager(doc_id).snapshot(
                        name, {"input": tool_input}
                    )
                    result["snapshot_id"] = snap_id
        return result


REGISTRY = ToolRegistry()
