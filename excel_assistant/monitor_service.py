from __future__ import annotations

import threading
from dataclasses import asdict
from datetime import date
from typing import Callable

from excel_assistant.config import SettingsStore
from excel_assistant.emailer import EmailSender
from excel_assistant.excel_monitor import ExcelMonitor
from excel_assistant.models import MonitoredEntry
from excel_assistant.notifications import DesktopNotifier
from excel_assistant.sent_registry import SentRegistry


class MonitorService:
    def __init__(
        self,
        settings_store: SettingsStore,
        on_results: Callable[[list[dict[str, object]]], None] | None = None,
    ) -> None:
        self._settings_store = settings_store
        self._monitor = ExcelMonitor()
        self._notifier = DesktopNotifier()
        self._emailer = EmailSender()
        self._registry = SentRegistry()
        self._on_results = on_results

        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._paused = False
        self._last_results: list[dict[str, object]] = []
        self._notified_today: set[str] = set()
        self._emailed_today: set[str] = set()

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._loop, name="excel-monitor", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=3)

    def pause(self) -> None:
        self._paused = True

    def resume(self) -> None:
        self._paused = False

    @property
    def paused(self) -> bool:
        return self._paused

    def get_last_results(self) -> list[dict[str, object]]:
        return self._last_results

    def run_once(self) -> list[dict[str, object]]:
        settings = self._settings_store.load()
        entries = self._monitor.scan(settings)
        sent_keys = self._registry.sent_cell_keys(settings.excel_file_path)

        for entry in entries:
            if entry.emailed:
                continue
            registry_key = f"{entry.sheet_name.strip().lower()}|{entry.cell.strip().upper()}"
            if registry_key in sent_keys:
                entry.emailed = True

        sent_entries: list[MonitoredEntry] = []
        for entry in entries:
            self._maybe_notify(entry)
            if entry.status == "due":
                if self._maybe_send_email(entry, settings):
                    sent_entries.append(entry)

        if sent_entries:
            self._registry.mark_sent_batch(
                settings.excel_file_path,
                sent_entries,
                requires_excel_sync=bool(settings.email_sent_column),
            )
            sent_keys = self._registry.sent_cell_keys(settings.excel_file_path)

        pending_excel_sync = [
            entry
            for entry in entries
            if f"{entry.sheet_name.strip().lower()}|{entry.cell.strip().upper()}" in sent_keys and not entry.emailed
        ]

        if pending_excel_sync:
            try:
                updated = self._monitor.mark_emailed(settings, [entry.row for entry in pending_excel_sync])
            except PermissionError:
                # Workbook is locked by Excel/OneDrive; emails are already sent,
                # but the sent flag cannot be written yet.
                updated = 0

            if updated:
                self._registry.mark_excel_synced_batch(settings.excel_file_path, pending_excel_sync)
                # Refresh status after write to keep UI state accurate.
                entries = self._monitor.scan(settings)
                sent_keys = self._registry.sent_cell_keys(settings.excel_file_path)
                for entry in entries:
                    if entry.emailed:
                        continue
                    registry_key = f"{entry.sheet_name.strip().lower()}|{entry.cell.strip().upper()}"
                    if registry_key in sent_keys:
                        entry.emailed = True

        serialized = [asdict(e) for e in entries]
        self._last_results = serialized
        if self._on_results:
            self._on_results(serialized)
        return serialized

    def _loop(self) -> None:
        while not self._stop_event.is_set():
            settings = self._settings_store.load()
            if not self._paused:
                try:
                    self.run_once()
                except Exception:
                    # Keep monitor alive even if one check fails due to file lock/network issues.
                    pass

            sleep_seconds = max(60, settings.poll_minutes * 60)
            self._stop_event.wait(sleep_seconds)

    def _maybe_notify(self, entry: MonitoredEntry) -> None:
        if entry.status == "good":
            return

        day_key = date.today().isoformat()
        dedupe_key = f"{entry.cell}:{entry.status}:{day_key}"
        if dedupe_key in self._notified_today:
            return

        self._notifier.send(
            title=f"Excel Assistant: {entry.status.upper()}",
            message=f"{entry.cell} is at {entry.days} banking days.",
        )
        self._notified_today.add(dedupe_key)

    def _maybe_send_email(self, entry: MonitoredEntry, settings) -> bool:
        if entry.emailed:
            return False
        if not settings.email.enabled:
            return False

        runtime_key = f"{settings.excel_file_path}:{entry.sheet_name}:{entry.cell}:{date.today().isoformat()}"
        if runtime_key in self._emailed_today:
            return False

        recipient = entry.recipient.strip()
        if not recipient:
            return False

        try:
            self._emailer.send_entry(settings.email, entry, recipient)
            self._emailed_today.add(runtime_key)
            return True
        except Exception:
            # SMTP errors should not crash monitor loop.
            return False
