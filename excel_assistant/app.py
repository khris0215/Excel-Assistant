from __future__ import annotations

from pathlib import Path

import webview

from excel_assistant.bridge import WebApiBridge
from excel_assistant.config import SettingsStore
from excel_assistant.monitor_service import MonitorService
from excel_assistant.tray import TrayController


class ExcelAssistantApp:
    def __init__(self) -> None:
        self._settings_store = SettingsStore()
        self._monitor = MonitorService(self._settings_store)
        self._bridge = WebApiBridge(self._settings_store, self._monitor)

        self._window: webview.Window | None = None
        self._tray: TrayController | None = None
        self._exit_requested = False

    def run(self, silent: bool = False) -> None:
        self._monitor.start()
        self._tray = TrayController(
            on_open_settings=self._show_window,
            on_run_now=lambda: self._monitor.run_once(),
            on_toggle_pause=self._toggle_pause,
            on_exit=self._shutdown,
            is_paused=lambda: self._monitor.paused,
        )
        self._tray.start()

        ui_path = Path(__file__).resolve().parent / "ui" / "index.html"
        self._window = webview.create_window(
            title="Excel Assistant",
            url=ui_path.as_uri(),
            js_api=self._bridge,
            width=1220,
            height=860,
            min_size=(1024, 720),
            hidden=silent,
            confirm_close=True,
        )

        webview.start(self._on_webview_ready)

    def _on_webview_ready(self) -> None:
        pass

    def _show_window(self) -> None:
        if not self._window:
            return
        try:
            self._window.show()
            self._window.restore()
        except Exception:
            return

    def _toggle_pause(self) -> None:
        if self._monitor.paused:
            self._monitor.resume()
            return
        self._monitor.pause()

    def _shutdown(self) -> None:
        if self._exit_requested:
            return
        self._exit_requested = True
        self._monitor.stop()
        if self._tray:
            self._tray.stop()
        if self._window:
            try:
                self._window.destroy()
            except Exception:
                return
