"""Runtime configuration. API keys come from the environment only — never from
source files."""
import os
from dataclasses import dataclass, field


@dataclass
class Budgets:
    """Anti-brute-force iteration limits (see agent/budget.py for the ladder)."""

    max_attempts_per_step: int = 4
    max_tool_calls_per_step: int = 8
    max_tool_calls_per_task: int = 40
    step_wall_clock_seconds: float = 240.0
    same_error_strategy_switch: int = 2  # identical signatures before forced switch


@dataclass
class Config:
    api_key: str = field(default_factory=lambda: os.environ.get("MINIMAX_API_KEY", ""))
    base_url: str = "https://api.minimaxi.com/anthropic"
    model: str = "MiniMax-M3"
    max_tokens: int = 8192
    visual_verify: bool = True
    max_verify_pages: int = 4
    render_timeout: float = 120.0
    non_interactive: bool = False
    auto_approve_overwrite: bool = False  # --yes
    verbose: bool = False
    budgets: Budgets = field(default_factory=Budgets)

    def require_api_key(self) -> None:
        if not self.api_key:
            raise RuntimeError(
                "MINIMAX_API_KEY is not set. Add it to your shell profile "
                "(e.g. ~/.zshrc) or a .env file — never hardcode it in source."
            )
