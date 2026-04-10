from __future__ import annotations

from typing import Any

from excel_assistant.config import SettingsStore
from excel_assistant.monitor_service import MonitorService
from excel_assistant.models import AppSettings
from excel_assistant.startup import set_startup


class WebApiBridge:
    def __init__(self, settings_store: SettingsStore, monitor: MonitorService) -> None:
        self._settings_store = settings_store
        self._monitor = monitor

    def get_settings(self) -> dict[str, Any]:
        return self._settings_store.load().to_dict()

    def save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        settings = AppSettings.from_dict(payload)
        self._settings_store.save(settings)
        try:
            set_startup(settings.autostart)
        except Exception:
            # Persisted settings should not fail just because startup registry update failed.
            pass
        return {"ok": True, "settings": settings.to_dict()}

    def run_check_now(self) -> dict[str, Any]:
        try:
            results = self._monitor.run_once()
            return {"ok": True, "results": results}
        except Exception as exc:
            return {"ok": False, "results": self._monitor.get_last_results(), "error": str(exc)}

    def get_last_results(self) -> list[dict[str, Any]]:
        return self._monitor.get_last_results()

    def pause_monitor(self) -> dict[str, Any]:
        self._monitor.pause()
        return {"ok": True, "paused": True}

    def resume_monitor(self) -> dict[str, Any]:
        self._monitor.resume()
        return {"ok": True, "paused": False}

    def monitor_state(self) -> dict[str, Any]:
        return {"paused": self._monitor.paused}
