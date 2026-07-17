"""Snapshot manager for document version control"""
import os
import shutil
import uuid
from typing import Dict, Optional
from pathlib import Path
import json
from datetime import datetime


class SnapshotManager:
    """Manages document snapshots for undo/redo functionality"""
    
    def __init__(self, snapshot_dir: str = ".snapshots"):
        self.snapshot_dir = Path(snapshot_dir)
        self.snapshot_dir.mkdir(exist_ok=True)
        self.snapshots: Dict[str, list] = {}  # doc_id -> list of snapshot_ids
    
    def create_snapshot(self, doc_id: str, source_path: str) -> str:
        """Create a snapshot of the document"""
        snapshot_id = str(uuid.uuid4())
        snapshot_path = self.snapshot_dir / doc_id / snapshot_id
        snapshot_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Copy file to snapshot
        shutil.copy2(source_path, str(snapshot_path))
        
        # Record metadata
        metadata = {
            "snapshot_id": snapshot_id,
            "doc_id": doc_id,
            "created_at": datetime.now().isoformat(),
            "source_path": source_path
        }
        metadata_path = snapshot_path.with_suffix(snapshot_path.suffix + ".meta")
        with open(metadata_path, 'w') as f:
            json.dump(metadata, f)
        
        # Track snapshot
        if doc_id not in self.snapshots:
            self.snapshots[doc_id] = []
        self.snapshots[doc_id].append(snapshot_id)
        
        return snapshot_id
    
    def restore_snapshot(self, doc_id: str, snapshot_id: str, target_path: str) -> bool:
        """Restore document from snapshot"""
        snapshot_path = self.snapshot_dir / doc_id / snapshot_id
        
        if not snapshot_path.exists():
            return False
        
        shutil.copy2(str(snapshot_path), target_path)
        return True
    
    def get_latest_snapshot(self, doc_id: str) -> Optional[str]:
        """Get the latest snapshot ID for a document"""
        if doc_id not in self.snapshots or not self.snapshots[doc_id]:
            return None
        return self.snapshots[doc_id][-1]
    
    def list_snapshots(self, doc_id: str) -> list:
        """List all snapshots for a document"""
        return self.snapshots.get(doc_id, [])
    
    def cleanup_old_snapshots(self, doc_id: str, keep_last_n: int = 10):
        """Clean up old snapshots, keeping only the last N"""
        if doc_id not in self.snapshots:
            return
        
        snapshots = self.snapshots[doc_id]
        if len(snapshots) <= keep_last_n:
            return
        
        # Remove old snapshots
        to_remove = snapshots[:-keep_last_n]
        for snapshot_id in to_remove:
            snapshot_path = self.snapshot_dir / doc_id / snapshot_id
            if snapshot_path.exists():
                snapshot_path.unlink()
            metadata_path = snapshot_path.with_suffix(snapshot_path.suffix + ".meta")
            if metadata_path.exists():
                metadata_path.unlink()
        
        # Update list
        self.snapshots[doc_id] = snapshots[-keep_last_n:]
