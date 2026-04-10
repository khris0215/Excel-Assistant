from __future__ import annotations

import json
from pathlib import Path

from excel_assistant.models import AppSettings


BASE_DIR = Path(__file__).resolve().parent
RUNTIME_DIR = BASE_DIR / "runtime"
SETTINGS_FILE = RUNTIME_DIR / "settings.json"


class SettingsStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or SETTINGS_FILE
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> AppSettings:
        if not self.path.exists():
            settings = AppSettings()
            self.save(settings)
            return settings

        with self.path.open("r", encoding="utf-8") as fh:
            raw = json.load(fh)
        return AppSettings.from_dict(raw)

    def save(self, settings: AppSettings) -> None:
        settings.normalize()
        payload = settings.to_dict()
        temp_path = self.path.with_suffix(".tmp")
        with temp_path.open("w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)
        temp_path.replace(self.path)
