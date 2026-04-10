from __future__ import annotations

from datetime import date, datetime
from typing import Iterable

from openpyxl.utils import column_index_from_string, get_column_letter


def business_days_between(entry: date, today: date) -> int:
    if entry > today:
        return 0

    # Count weekdays including both endpoints, then subtract one to get elapsed banking days.
    total_days = (today - entry).days + 1
    full_weeks, extra_days = divmod(total_days, 7)
    weekdays = full_weeks * 5

    start_weekday = entry.weekday()
    for offset in range(extra_days):
        weekday = (start_weekday + offset) % 7
        if weekday < 5:
            weekdays += 1

    return max(0, weekdays - 1)


def normalize_col(col: str) -> str:
    return (col or "").strip().upper()


def parse_list_csv(value: str) -> list[str]:
    return [item.strip() for item in (value or "").split(",") if item.strip()]


def to_date(value: object) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        # Avoid interpreting plain numbers (e.g., IDs/status codes) as epoch-based dates.
        return None

    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None

        # Fast-path ISO date/time values first.
        try:
            return date.fromisoformat(text)
        except ValueError:
            pass

        try:
            return datetime.fromisoformat(text).date()
        except ValueError:
            pass

        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue

        return None

    return None


def col_to_index(col: str) -> int:
    return column_index_from_string(normalize_col(col))


def index_to_col(index: int) -> str:
    return get_column_letter(index)


def dedupe_rows(rows: Iterable[int]) -> list[int]:
    return sorted(set(rows))
