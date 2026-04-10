from __future__ import annotations

from plyer import notification


class DesktopNotifier:
    def send(self, title: str, message: str, timeout: int = 8) -> None:
        notification.notify(
            title=title,
            message=message,
            app_name="Excel Assistant",
            timeout=timeout,
        )
