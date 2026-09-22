"""Project paths, .env configuration and SMTP connection handling."""

from __future__ import annotations

import os
import smtplib
import socket
import ssl
from dataclasses import dataclass
from pathlib import Path

from dotenv import dotenv_values

BASE_DIR = Path(__file__).resolve().parent
ENV_FILE = BASE_DIR / ".env"
DOCX_PATH = BASE_DIR / "emails.docx"
ATTACHMENTS_DIR = BASE_DIR / "attachments"
DATA_DIR = BASE_DIR / "data"
QUEUE_CSV = DATA_DIR / "queue.csv"
QUEUE_BACKUP_CSV = DATA_DIR / "queue.backup.csv"
LOG_FILE = DATA_DIR / "activity.log"

QUEUE_COLUMNS = [
    "id",
    "organization",
    "recipient_email",
    "subject",
    "body",
    "lang",
    "attachment_filename",
    "status",
    "error_message",
    "sent_at",
]
STATUS_PENDING, STATUS_SENT, STATUS_ERROR = "pending", "sent", "error"
STATUSES = [STATUS_PENDING, STATUS_SENT, STATUS_ERROR]
LANGS = ["ua", "en"]

# Used when a letter does not name its attachment explicitly.
DEFAULT_ATTACHMENTS = {
    "ua": "expense_estimate_ua.pdf",
    "en": "expense_estimate_en.pdf",
}
# Put this in attachment_filename to send a letter without any attachment.
NO_ATTACHMENT = "none"

SECURITY_MODES = ("ssl", "starttls", "none")
REQUIRED_ENV = ("SMTP_SERVER", "SMTP_PORT", "SENDER_EMAIL", "SENDER_PASSWORD")


class ConfigError(Exception):
    """Raised when .env is missing required SMTP settings."""


@dataclass(frozen=True)
class SmtpSettings:
    server: str
    port: int
    sender_email: str
    password: str
    username: str
    sender_name: str
    security: str
    reply_to: str
    debug_email: str
    timeout: int

    def missing_fields(self) -> list[str]:
        values = {
            "SMTP_SERVER": self.server,
            "SMTP_PORT": self.port,
            "SENDER_EMAIL": self.sender_email,
            "SENDER_PASSWORD": self.password,
        }
        return [name for name in REQUIRED_ENV if not values[name]]

    @property
    def is_complete(self) -> bool:
        return not self.missing_fields()

    def describe(self) -> str:
        return f"{self.server or '?'}:{self.port or '?'} ({self.security.upper()})"


_ENV_KEYS = {
    "SMTP_SERVER",
    "SMTP_PORT",
    "SMTP_SECURITY",
    "SMTP_USERNAME",
    "SMTP_TIMEOUT",
    "SENDER_EMAIL",
    "SENDER_PASSWORD",
    "SENDER_NAME",
    "REPLY_TO",
    "DEBUG_EMAIL",
}


def _env() -> dict[str, str]:
    """Values from .env take precedence over real environment variables.

    The file is re-read on every call, so edits to .env apply without
    restarting Streamlit.
    """
    merged = {k: v for k, v in os.environ.items() if k in _ENV_KEYS}
    if ENV_FILE.exists():
        merged.update({k: v for k, v in dotenv_values(ENV_FILE).items() if v is not None})
    return {k: v.strip() for k, v in merged.items()}


def load_settings() -> SmtpSettings:
    env = _env()
    try:
        port = int(env.get("SMTP_PORT") or 0)
    except ValueError:
        port = 0
    try:
        timeout = int(env.get("SMTP_TIMEOUT") or 30)
    except ValueError:
        timeout = 30

    security = (env.get("SMTP_SECURITY") or "auto").lower()
    if security not in SECURITY_MODES:
        # auto: implicit TLS on 465, STARTTLS everywhere else (587, 25, 2525).
        security = "ssl" if port == 465 else "starttls"

    server = env.get("SMTP_SERVER", "")
    password = env.get("SENDER_PASSWORD", "")
    if "gmail" in server.lower():
        # Google shows app passwords as "abcd efgh ijkl mnop"; the spaces are not part of it.
        password = password.replace(" ", "")

    sender_email = env.get("SENDER_EMAIL", "")
    return SmtpSettings(
        server=server,
        port=port,
        sender_email=sender_email,
        password=password,
        username=env.get("SMTP_USERNAME") or sender_email,
        sender_name=env.get("SENDER_NAME", ""),
        security=security,
        reply_to=env.get("REPLY_TO", ""),
        debug_email=env.get("DEBUG_EMAIL", ""),
        timeout=timeout,
    )


def open_smtp(settings: SmtpSettings) -> smtplib.SMTP:
    """Return a connected, authenticated SMTP client (usable as a context manager)."""
    missing = settings.missing_fields()
    if missing:
        raise ConfigError(f"Missing in .env: {', '.join(missing)}")

    context = ssl.create_default_context()
    if settings.security == "ssl":
        smtp: smtplib.SMTP = smtplib.SMTP_SSL(
            settings.server, settings.port, timeout=settings.timeout, context=context
        )
    else:
        smtp = smtplib.SMTP(settings.server, settings.port, timeout=settings.timeout)
    try:
        smtp.ehlo()
        if settings.security == "starttls":
            smtp.starttls(context=context)
            smtp.ehlo()
        smtp.login(settings.username, settings.password)
    except BaseException:
        smtp.close()
        raise
    return smtp


def test_connection(settings: SmtpSettings | None = None) -> tuple[bool, str]:
    settings = settings or load_settings()
    try:
        with open_smtp(settings) as smtp:
            smtp.noop()
    except Exception as exc:  # noqa: BLE001 - every failure is reported to the UI
        return False, describe_smtp_error(exc, settings)
    return True, f"Connected and logged in to {settings.describe()} as {settings.username}"


def describe_smtp_error(exc: BaseException, settings: SmtpSettings | None = None) -> str:
    """Turn an SMTP/socket exception into a short, actionable message."""
    hint = ""
    if isinstance(exc, ConfigError):
        return f"{exc}. Copy .env.template to .env and fill it in."
    if isinstance(exc, smtplib.SMTPAuthenticationError):
        hint = (
            "Authentication failed: check SENDER_EMAIL / SENDER_PASSWORD. "
            "Gmail, Outlook and Ukr.net need an app password, not your normal password."
        )
    elif isinstance(exc, smtplib.SMTPRecipientsRefused):
        refused = ", ".join(f"{addr} ({code} {_decode(msg)})" for addr, (code, msg) in exc.recipients.items())
        return f"Recipient refused: {refused}"
    elif isinstance(exc, smtplib.SMTPResponseException):
        return f"SMTP {exc.smtp_code}: {_decode(exc.smtp_error)}"
    elif isinstance(exc, socket.gaierror):
        hint = f"Unknown SMTP host '{settings.server if settings else ''}'. Check SMTP_SERVER."
    elif isinstance(exc, ssl.SSLError):
        hint = "TLS handshake failed. Port 465 needs SSL; port 587 needs STARTTLS (SMTP_SECURITY=auto picks this)."
    elif isinstance(exc, (TimeoutError, socket.timeout)):
        hint = "Connection timed out. Check SMTP_SERVER/SMTP_PORT, your network, and that the port matches the security mode."
    elif isinstance(exc, ConnectionRefusedError):
        hint = "Connection refused. Check SMTP_PORT."
    elif isinstance(exc, smtplib.SMTPServerDisconnected):
        hint = "Server closed the connection. Often a port/security mismatch (465 = SSL, 587 = STARTTLS)."
    detail = f"{type(exc).__name__}: {exc}" if str(exc) else type(exc).__name__
    return f"{hint} ({detail})" if hint else detail


def _decode(value: bytes | str) -> str:
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return str(value)
