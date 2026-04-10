from __future__ import annotations

import threading
from typing import Callable

from PIL import Image, ImageDraw
import pystray


class TrayController:
    def __init__(
        self,
        on_open_settings: Callable[[], None],
        on_run_now: Callable[[], None],
        on_toggle_pause: Callable[[], None],
        on_exit: Callable[[], None],
        is_paused: Callable[[], bool],
    ) -> None:
        self._on_open_settings = on_open_settings
        self._on_run_now = on_run_now
        self._on_toggle_pause = on_toggle_pause
        self._on_exit = on_exit
        self._is_paused = is_paused

        self.icon = pystray.Icon("ExcelAssistant", self._build_icon(), "Excel Assistant", self._menu())
        self._thread: threading.Thread | None = None

    def _build_icon(self) -> Image.Image:
        img = Image.new("RGBA", (64, 64), (16, 32, 56, 255))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle((8, 8, 56, 56), radius=10, fill=(23, 162, 122, 255))
        draw.rectangle((18, 18, 46, 24), fill=(255, 255, 255, 220))
        draw.rectangle((18, 29, 46, 35), fill=(255, 255, 255, 220))
        draw.rectangle((18, 40, 46, 46), fill=(255, 255, 255, 220))
        return img

    def _pause_label(self, *_args) -> str:
        return "Resume Monitoring" if self._is_paused() else "Pause Monitoring"

    def _menu(self):
        return pystray.Menu(
            pystray.MenuItem("Open Dashboard", lambda icon, item: self._on_open_settings()),
            pystray.MenuItem("Run Check Now", lambda icon, item: self._on_run_now()),
            pystray.MenuItem(self._pause_label, lambda icon, item: self._on_toggle_pause()),
            pystray.MenuItem("Exit", lambda icon, item: self._exit(icon)),
        )

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self.icon.run, daemon=True, name="tray")
        self._thread.start()

    def stop(self) -> None:
        self.icon.stop()

    def _exit(self, icon) -> None:
        self._on_exit()
        icon.stop()
