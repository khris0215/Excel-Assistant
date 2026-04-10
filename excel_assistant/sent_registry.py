from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock

from excel_assistant.config import RUNTIME_DIR


REGISTRY_FILE = RUNTIME_DIR / "sent_registry.json"


class SentRegistry:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or REGISTRY_FILE
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = Lock()

    def was_sent(self, workbook_path: str, sheet_name: str, cell: str) -> bool:
        key = self._key(workbook_path, sheet_name, cell)
        data = self._load()
        record = data["records"].get(key, {})
        return bool(record.get("sent", False))

    def sent_cell_keys(self, workbook_path: str) -> set[str]:
        data = self._load()
        normalized_path = self._normalized_path(workbook_path)
        prefix = f"{normalized_path}|"
        sent: set[str] = set()
        for key, record in data["records"].items():
            if not key.startswith(prefix):
                continue
            if not isinstance(record, dict) or not record.get("sent"):
                continue
            _, sheet_name, cell = key.split("|", 2)
            sent.add(f"{sheet_name}|{cell}")
        return sent

    def mark_sent_batch(self, workbook_path: str, entries, requires_excel_sync: bool) -> None:
        now = datetime.now(timezone.utc).isoformat()
        with self._lock:
            data = self._load_unlocked()

            for entry in entries:
                key = self._key(workbook_path, entry.sheet_name, entry.cell)
                data["records"][key] = {
                    "sent": True,
                    "workbook_path": str(Path(workbook_path).expanduser()),
                    "sheet_name": entry.sheet_name,
                    "cell": entry.cell,
                    "row": entry.row,
                    "recipient": entry.recipient,
                    "status": entry.status,
                    "days": entry.days,
                    "sent_at": now,
                    "excel_synced": not requires_excel_sync,
                }

            self._save_unlocked(data)

    def mark_excel_synced_batch(self, workbook_path: str, entries) -> None:
        with self._lock:
            data = self._load_unlocked()
            for entry in entries:
                key = self._key(workbook_path, entry.sheet_name, entry.cell)
                record = data["records"].get(key)
                if not record:
                    continue
                record["excel_synced"] = True
            self._save_unlocked(data)

    def _load(self) -> dict[str, object]:
        with self._lock:
            return self._load_unlocked()

    def _save(self, payload: dict[str, object]) -> None:
        with self._lock:
            self._save_unlocked(payload)

    def _load_unlocked(self) -> dict[str, object]:
        if not self.path.exists():
            return {"version": 1, "records": {}}

        try:
            with self.path.open("r", encoding="utf-8") as fh:
                raw = json.load(fh)
        except (OSError, json.JSONDecodeError):
            return {"version": 1, "records": {}}

        records = raw.get("records", {}) if isinstance(raw, dict) else {}
        if not isinstance(records, dict):
            records = {}
        return {"version": 1, "records": records}

    def _save_unlocked(self, payload: dict[str, object]) -> None:
        temp_path = self.path.with_suffix(".tmp")
        with temp_path.open("w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)
        temp_path.replace(self.path)

    @staticmethod
    def _key(workbook_path: str, sheet_name: str, cell: str) -> str:
        normalized_path = SentRegistry._normalized_path(workbook_path)
        return f"{normalized_path}|{(sheet_name or '').strip().lower()}|{(cell or '').strip().upper()}"

    @staticmethod
    def _normalized_path(workbook_path: str) -> str:
        path = Path(workbook_path).expanduser()
        try:
            return str(path.resolve()).lower()
        except OSError:
            return str(path.absolute()).lower()
