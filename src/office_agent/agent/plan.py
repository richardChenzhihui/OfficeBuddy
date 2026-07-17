"""Plan model: the step checklist shown to the user and tracked by the loop."""
from dataclasses import dataclass, field
from typing import List, Optional

STATUSES = ("todo", "in_progress", "done", "blocked")


@dataclass
class PlanStep:
    description: str
    status: str = "todo"


@dataclass
class Plan:
    steps: List[PlanStep] = field(default_factory=list)
    active_index: Optional[int] = None  # most recently activated step

    @classmethod
    def from_descriptions(cls, descriptions: List[str]) -> "Plan":
        return cls(steps=[PlanStep(d) for d in descriptions])

    def set_status(self, index: int, status: str) -> None:
        if not 0 <= index < len(self.steps):
            raise ValueError(
                f"Step index {index} out of range: plan has {len(self.steps)} steps."
            )
        if status not in STATUSES:
            raise ValueError(f"Invalid status '{status}': expected one of {STATUSES}.")
        self.steps[index].status = status
        if status == "in_progress":
            self.active_index = index

    @property
    def current_index(self) -> int:
        # Prefer the most recently activated step: a step stuck 'in_progress'
        # after a failed verification must not hijack budget attribution for
        # work the model is doing on a later step.
        if (
            self.active_index is not None
            and 0 <= self.active_index < len(self.steps)
            and self.steps[self.active_index].status == "in_progress"
        ):
            return self.active_index
        for i, step in enumerate(self.steps):
            if step.status == "in_progress":
                return i
        for i, step in enumerate(self.steps):
            if step.status == "todo":
                return i
        return max(len(self.steps) - 1, 0)

    @property
    def all_done(self) -> bool:
        return bool(self.steps) and all(s.status == "done" for s in self.steps)

    def render_text(self) -> str:
        icons = {"todo": "☐", "in_progress": "⏳", "done": "✅", "blocked": "⚠"}
        return "\n".join(
            f"{icons[s.status]} {i}. {s.description}" for i, s in enumerate(self.steps)
        )
