from __future__ import annotations

import smtplib
from email.message import EmailMessage

from excel_assistant.models import EmailSettings, MonitoredEntry


class EmailSender:
    def send_entry(self, cfg: EmailSettings, entry: MonitoredEntry, recipient: str) -> None:
        message = EmailMessage()
        message["Subject"] = cfg.subject_template.format(
            row=entry.row,
            cell=entry.cell,
            days=entry.days,
            entry_date=entry.entry_date,
            status=entry.status,
            recipient=recipient,
        )
        message["From"] = cfg.sender_email or cfg.smtp_username
        message["To"] = recipient
        message.set_content(
            cfg.body_template.format(
                row=entry.row,
                cell=entry.cell,
                days=entry.days,
                entry_date=entry.entry_date,
                status=entry.status,
                recipient=recipient,
            )
        )

        with smtplib.SMTP(cfg.smtp_host, cfg.smtp_port, timeout=20) as smtp:
            if cfg.use_tls:
                smtp.starttls()
            if cfg.smtp_username:
                smtp.login(cfg.smtp_username, cfg.smtp_password)
            smtp.send_message(message)
