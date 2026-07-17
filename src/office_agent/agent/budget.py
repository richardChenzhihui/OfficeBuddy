"""Anti-brute-force iteration budgets and the escalation ladder.

Ladder per plan step:
  attempt 1-2 (same signature)  -> targeted retry with the concrete error detail
  same signature twice in a row -> forced strategy switch instruction
  failure after strategy switch -> ask the user (multiple choice)
  hard cap reached              -> ask the user, never silently continue
"""
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, List

from ..config import Budgets


class Action(str, Enum):
    RETRY = "retry"
    SWITCH_STRATEGY = "switch_strategy"
    ASK_USER = "ask_user"


_CATEGORIES = [
    ("no_match", re.compile(r"matched no elements|not found|no paragraph|0 match", re.I)),
    ("out_of_range", re.compile(r"out of range", re.I)),
    ("invalid_args", re.compile(r"invalid (arguments|excel range|chart|color|alignment)", re.I)),
    ("verify_text", re.compile(r"text|content|wording|missing|duplicat", re.I)),
    ("verify_style", re.compile(r"font|bold|italic|size|color|align|format|style", re.I)),
    ("verify_layout", re.compile(r"layout|overlap|page|position|spacing", re.I)),
]


def normalize_signature(source: str, detail: str) -> str:
    """Collapse an error/verification failure to a coarse signature so that
    'the same kind of wrong' matches even when free-text details differ.
    source: tool name or 'verify'."""
    for category, pattern in _CATEGORIES:
        if pattern.search(detail):
            return f"{source}:{category}"
    return f"{source}:other"


@dataclass
class StepBudget:
    attempts: int = 0
    tool_calls: int = 0
    signatures: List[str] = field(default_factory=list)
    strategy_switched: bool = False
    user_asks: int = 0


@dataclass
class BudgetTracker:
    limits: Budgets
    steps: Dict[int, StepBudget] = field(default_factory=dict)
    total_tool_calls: int = 0
    end_turn_repairs: int = 0

    def step(self, index: int) -> StepBudget:
        return self.steps.setdefault(index, StepBudget())

    def record_tool_call(self, step_index: int) -> None:
        self.total_tool_calls += 1
        self.step(step_index).tool_calls += 1

    def task_exhausted(self) -> bool:
        return self.total_tool_calls >= self.limits.max_tool_calls_per_task

    def step_exhausted(self, step_index: int) -> bool:
        """A single step spinning on tool calls without converging."""
        return self.step(step_index).tool_calls >= self.limits.max_tool_calls_per_step

    def record_end_turn_repair(self) -> None:
        self.end_turn_repairs += 1

    def end_turn_exhausted(self) -> bool:
        """Hard cap on the end-turn verify->repair cycle, independent of tool
        calls, so a model that stops calling tools cannot loop forever."""
        return self.end_turn_repairs >= self.limits.max_attempts_per_step

    def reset_steps(self) -> None:
        """New plan: per-step budgets restart; task-level counters persist."""
        self.steps.clear()

    def record_failure(self, step_index: int, signature: str) -> Action:
        budget = self.step(step_index)
        budget.attempts += 1
        budget.signatures.append(signature)

        if budget.attempts >= self.limits.max_attempts_per_step:
            return Action.ASK_USER
        if budget.strategy_switched:
            # Strategy switch already tried and it still failed.
            return Action.ASK_USER
        recent = budget.signatures[-self.limits.same_error_strategy_switch :]
        if (
            len(recent) == self.limits.same_error_strategy_switch
            and len(set(recent)) == 1
        ):
            budget.strategy_switched = True
            return Action.SWITCH_STRATEGY
        return Action.RETRY
