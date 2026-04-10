from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class Thresholds:
    good_max: int = 20
    soft_max: int = 28
    medium_max: int = 35
    hard_max: int = 44
    due_at: int = 45

    def normalize(self) -> None:
        values = sorted([
            max(0, int(self.good_max)),
            max(0, int(self.soft_max)),
            max(0, int(self.medium_max)),
            max(0, int(self.hard_max)),
            max(1, int(self.due_at)),
        ])
        self.good_max, self.soft_max, self.medium_max, self.hard_max, self.due_at = values


@dataclass
class WatchSelection:
    mode: str = "range"  # range | rows_all_columns | columns_all_rows
    sheet_name: str = ""
    start_row: int = 2
    end_row: int = 200
    start_col: str = "A"
    end_col: str = "A"
    row_list: str = ""
    column_list: str = ""


@dataclass
class EmailSettings:
    enabled: bool = False
    recipient_mode: str = "global"  # global | excel_column
    global_recipient: str = ""
    email_column: str = ""
    sender_email: str = ""
    smtp_host: str = ""
    smtp_port: int = 587
    smtp_username: str = ""
    smtp_password: str = ""
    use_tls: bool = True
    subject_template: str = "[Excel Assistant] {status} item at {cell}"
    body_template: str = (
        "Hello,\\n\\n"
        "Item details:\\n"
        "- Cell: {cell}\\n"
        "- Row: {row}\\n"
        "- Entry Date: {entry_date}\\n"
        "- Banking Days: {days}\\n"
        "- Status: {status}\\n\\n"
        "Regards,\\nExcel Assistant"
    )


@dataclass
class AppSettings:
    excel_file_path: str = ""
    poll_minutes: int = 5
    autostart: bool = False
    watch: WatchSelection = field(default_factory=WatchSelection)
    thresholds: Thresholds = field(default_factory=Thresholds)
    email_sent_column: str = ""
    email: EmailSettings = field(default_factory=EmailSettings)

    def normalize(self) -> None:
        self.poll_minutes = max(1, int(self.poll_minutes))
        self.thresholds.normalize()
        self.watch.start_row = max(1, int(self.watch.start_row))
        self.watch.end_row = max(self.watch.start_row, int(self.watch.end_row))

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AppSettings":
        watch = WatchSelection(**data.get("watch", {}))
        thresholds = Thresholds(**data.get("thresholds", {}))
        email = EmailSettings(**data.get("email", {}))
        settings = cls(
            excel_file_path=data.get("excel_file_path", ""),
            poll_minutes=data.get("poll_minutes", 5),
            autostart=data.get("autostart", False),
            watch=watch,
            thresholds=thresholds,
            email_sent_column=data.get("email_sent_column", ""),
            email=email,
        )
        settings.normalize()
        return settings


@dataclass
class MonitoredEntry:
    sheet_name: str
    row: int
    column: str
    cell: str
    entry_date: str
    days: int
    status: str
    recipient: str
    emailed: bool
