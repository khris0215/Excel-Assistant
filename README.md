# Excel Assistant

A Windows desktop tray application (Python backend + HTML/Tailwind/JavaScript UI) that monitors a single Excel workbook and alerts users as banking-day age thresholds are reached.

## What This Version Delivers

- Beautiful, web-style desktop UI (no tkinter), powered by `pywebview`.
- Windows tray app (`pystray`) that can stay in background.
- Startup on login support via Windows registry Run key.
- Configurable monitoring scope:
  - Specific row + column range
  - Specific rows + all columns
  - Specific columns + all rows in range
- Banking day age calculation (weekdays only).
- Configurable thresholds:
  - Good
  - Soft
  - Medium
  - Hard
  - Due
- Desktop notifications (`plyer`) for non-good statuses.
- Automated SMTP emails (`smtplib`) when entries are Due.
- Per-row recipient from Excel column, or global recipient from settings.
- Duplicate prevention via configurable `Email Sent` column marker.
- Modular code organization for maintainability.

## Project Structure

```text
Excel-Assistant/
  main.py
  requirements.txt
  README.md
  excel_assistant/
    __init__.py
    app.py
    bridge.py
    config.py
    emailer.py
    excel_monitor.py
    models.py
    monitor_service.py
    notifications.py
    startup.py
    tray.py
    utils.py
    runtime/
      .gitkeep
      settings.json               # auto-created at first run
    ui/
      index.html
      app.js
```

## Installation

1. Create and activate virtual environment.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install dependencies.

```powershell
pip install -r requirements.txt
```

## Run

Standard launch:

```powershell
python main.py
```

Start hidden/minimized for startup use:

```powershell
python main.py --silent
```

## Settings Guide

### Monitoring Scope

- `Watch Mode = Specific row + column range`:
  - Set `start_row`, `end_row`, `start_col`, `end_col`
- `Watch Mode = Specific rows + all columns`:
  - Set `row_list` like `2,4,8,20`
- `Watch Mode = Specific columns + all rows in range`:
  - Set `column_list` like `C,F,H`, and row bounds

### Thresholds

Set the boundaries in banking days:

- `good_max`
- `soft_max`
- `medium_max`
- `hard_max`
- `due_at`

Default behavior:

- `<= good_max`: Good
- `<= soft_max`: Soft
- `<= medium_max`: Medium
- `<= hard_max`: Hard
- `>= due_at`: Due

### Email

- Enable `email.enabled`
- Choose `recipient_mode`:
  - `global`: sends to one configured recipient
  - `excel_column`: reads recipient from row email column
- Configure SMTP host/port/credentials, sender, subject, and body templates.
- Configure `email_sent_column` so sent rows are marked `Yes`.

Template placeholders supported:

- `{row}`
- `{cell}`
- `{days}`
- `{entry_date}`
- `{status}`
- `{recipient}`

## How It Works

1. Background monitor loop runs every `poll_minutes`.
2. Workbook is read with `openpyxl` in read-only mode.
3. Each watched cell is parsed for date values.
4. Banking days are computed from entry date to today.
5. Non-good statuses trigger desktop notifications.
6. Due entries trigger emails (if enabled), then row is marked as emailed.

## Packaging to EXE (Later)

You can package this to a single executable with PyInstaller.

1. Install PyInstaller:

```powershell
pip install pyinstaller
```

2. Build command:

```powershell
pyinstaller --noconfirm --onefile --windowed --name ExcelAssistant --add-data "excel_assistant/ui;excel_assistant/ui" main.py
```

Note: For tray/background behavior in production, many teams prefer `--onedir` over `--onefile` for easier static-asset handling.

## Security Notes

- SMTP password is currently stored in plaintext JSON settings.
- Add encryption/credential vault integration before production rollout.

## Next Suggested Improvements

1. Add holiday calendar support in banking-day calculation.
2. Add in-app logs panel with failures and retry tracking.
3. Add test suite for date classification and watch-mode parsing.
4. Add richer templating with optional row context fields.

How to run

Create virtual env: python -m venv .venv
Activate: .\.venv\Scripts\Activate.ps1
Install deps: pip install -r requirements.txt
Launch app: python main.py
Optional tray-minimized start: python main.py --silent



link to register app password in google account for auto email
https://myaccount.google.com/apppasswords

make sure 2fa is on