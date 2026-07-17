"""Rich + questionary terminal UI implementing the loop's UI protocol."""
import json
from typing import Any, Dict

from rich.console import Console
from rich.markdown import Markdown
from rich.panel import Panel

from ..agent.loop import BaseUI
from ..agent.plan import Plan
from ..tools.interaction_tools import AskUserInput

_ICONS = {"todo": "☐", "in_progress": "⏳", "done": "✅", "blocked": "⚠️"}


class CliUI(BaseUI):
    def __init__(self, verbose: bool = False, non_interactive: bool = False):
        self.console = Console()
        self.verbose = verbose
        self.non_interactive = non_interactive

    def emit_text(self, text: str) -> None:
        self.console.print(Markdown(text))

    def notify(self, message: str) -> None:
        self.console.print(f"[dim]· {message}[/dim]")

    def tool_call(self, name: str, tool_input: Dict[str, Any]) -> None:
        if self.verbose:
            args = json.dumps(tool_input, ensure_ascii=False, default=str)
            if len(args) > 160:
                args = args[:160] + "…"
            self.console.print(f"[cyan]→ {name}[/cyan][dim]({args})[/dim]")
        else:
            self.console.print(f"[cyan]→ {name}[/cyan]")

    def tool_result(self, name: str, result: Dict[str, Any]) -> None:
        if not result.get("success", True):
            self.console.print(f"[red]  ✗ {result.get('error', 'failed')}[/red]")
        elif self.verbose:
            brief = json.dumps(result, ensure_ascii=False, default=str)
            if len(brief) > 200:
                brief = brief[:200] + "…"
            self.console.print(f"[dim]  ✓ {brief}[/dim]")

    def plan_update(self, plan: Plan) -> None:
        lines = [
            f"{_ICONS[s.status]} {i}. {s.description}"
            for i, s in enumerate(plan.steps)
        ]
        self.console.print(Panel("\n".join(lines), title="计划", border_style="blue"))

    def ask_user(self, params: AskUserInput) -> Dict[str, Any]:
        if self.non_interactive:
            return super().ask_user(params)
        import questionary

        self.console.print()
        if params.kind == "multiple_choice" and params.options:
            choices = [
                questionary.Choice(
                    title=(
                        f"{o.label} — {o.description}" if o.description else o.label
                    ),
                    value=o.label,
                )
                for o in params.options
            ]
            choices.append(questionary.Choice(title="其他（自由输入）", value="__other__"))
            if params.allow_multiple:
                answer = questionary.checkbox(params.question, choices=choices).ask()
            else:
                answer = questionary.select(params.question, choices=choices).ask()
            if answer == "__other__" or (
                isinstance(answer, list) and "__other__" in answer
            ):
                answer = questionary.text("请输入：").ask()
        else:
            answer = questionary.text(params.question).ask()
        if answer is None:  # user hit Ctrl-C on the prompt
            return {"answer": "", "note": "User cancelled the question."}
        return {"answer": answer}
