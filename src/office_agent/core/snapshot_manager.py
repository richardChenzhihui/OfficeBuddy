"""Per-step snapshots of the in-memory document, with a persisted index.

Snapshots capture the in-memory object (serialized bytes) immediately after
each mutating tool call — not the on-disk file — so undo works between saves.
The index is persisted to index.json so snapshots survive process restarts.
"""
import json
import time
from pathlib import Path
from typing import Any, Dict, List

from .session import EditSession


class SnapshotManager:
    def __init__(self, session: EditSession):
        self.session = session
        self.dir = session.session_dir / "snapshots"
        self.dir.mkdir(exist_ok=True)
        self.index_path = self.dir / "index.json"
        self._index: List[Dict[str, Any]] = self._load_index()
        if not self._index:
            # Baseline: state as opened, so undo can always reach the start.
            self.snapshot("baseline", {})

    def _load_index(self) -> List[Dict[str, Any]]:
        if self.index_path.exists():
            return json.loads(self.index_path.read_text(encoding="utf-8"))
        return []

    def _save_index(self) -> None:
        self.index_path.write_text(
            json.dumps(self._index, ensure_ascii=False, indent=1), encoding="utf-8"
        )

    def snapshot(self, label: str, detail: Dict[str, Any]) -> str:
        seq = len(self._index)
        ext = self.session.working_path.suffix
        snap_id = f"{seq:03d}__{label}"
        path = self.dir / f"{snap_id}{ext}"
        path.write_bytes(self.session.to_bytes())
        self._index.append(
            {
                "id": snap_id,
                "seq": seq,
                "label": label,
                "ts": time.time(),
                "file": path.name,
                "detail": detail,
            }
        )
        self._save_index()
        return snap_id

    def list(self) -> List[Dict[str, Any]]:
        return [
            {k: entry[k] for k in ("id", "seq", "label", "ts")} for entry in self._index
        ]

    def restore(self, snap_id: str) -> Dict[str, Any]:
        entry = next((e for e in self._index if e["id"] == snap_id), None)
        if entry is None:
            raise KeyError(
                f"Unknown snapshot '{snap_id}': available are "
                f"{[e['id'] for e in self._index]}."
            )
        data = (self.dir / entry["file"]).read_bytes()
        self.session.reload_from_bytes(data)
        # Truncate history after the restored point so redo states don't linger.
        keep = entry["seq"] + 1
        for stale in self._index[keep:]:
            (self.dir / stale["file"]).unlink(missing_ok=True)
        self._index = self._index[:keep]
        self._save_index()
        return {"restored_to": snap_id, "label": entry["label"]}

    def undo(self, steps: int = 1) -> Dict[str, Any]:
        if steps < 1:
            raise ValueError("steps must be >= 1.")
        if len(self._index) <= 1:
            raise ValueError("Nothing to undo: no edits have been made yet.")
        target_seq = max(0, len(self._index) - 1 - steps)
        return self.restore(self._index[target_seq]["id"])
