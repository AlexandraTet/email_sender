"""Queue storage, message building and the background dispatch worker."""

from __future__ import annotations

import html
import mimetypes
import os
import random
import re
import smtplib
import threading
import time
from collections import Counter, deque
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

import pandas as pd

import config
from config import (
    ATTACHMENTS_DIR,
    DEFAULT_ATTACHMENTS,
    LOG_FILE,
    NO_ATTACHMENT,
    QUEUE_COLUMNS,
    STATUS_ERROR,
    STATUS_PENDING,
    STATUS_SENT,
    STATUSES,
)

EMAIL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}$")
LINK_RE = re.compile(
    r"(?P<url>https?://[^\s<>\"]+[^\s<>\".,;:!?)\]»'])"
    r"|(?P<email>[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,})"
)
MIN_DELAY_SECONDS = 5
RETRY_BACKOFF_SECONDS = (30, 90)  # waits before 2nd and 3rd attempt on connection errors


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# --------------------------------------------------------------------------- queue storage


class QueueLockedError(Exception):
    """queue.csv could not be written, usually because it is open in Excel."""


class QueueStore:
    """Thread-safe access to data/queue.csv.

    The Streamlit sessions and the worker thread live in one process, so a
    process-wide lock is enough. Every write re-reads the file first and only
    touches the rows it means to change.
    """

    def __init__(self, path: Path):
        self.path = path
        self.lock = threading.RLock()  # hold it for multi-step read-modify-write
        self._cache: tuple[int, pd.DataFrame] | None = None

    def exists(self) -> bool:
        return self.path.exists()

    def mtime(self) -> int:
        try:
            return self.path.stat().st_mtime_ns
        except FileNotFoundError:
            return 0

    def load(self) -> pd.DataFrame:
        with self.lock:
            if not self.path.exists():
                return normalize_queue(pd.DataFrame(columns=QUEUE_COLUMNS))
            mtime = self.mtime()
            if self._cache is None or self._cache[0] != mtime:
                df = pd.read_csv(self.path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
                self._cache = (mtime, normalize_queue(df))
            return self._cache[1].copy()

    def save(self, df: pd.DataFrame) -> None:
        df = normalize_queue(df)
        with self.lock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.path.with_name(self.path.name + ".tmp")
            df.to_csv(tmp, index=False, encoding="utf-8-sig")
            for attempt in range(5):
                try:
                    os.replace(tmp, self.path)
                    break
                except PermissionError:
                    if attempt == 4:
                        tmp.unlink(missing_ok=True)
                        raise QueueLockedError(
                            f"Cannot write {self.path.name}: close it in Excel or any other program."
                        ) from None
                    time.sleep(0.3)
            self._cache = None

    def update_row(self, row_id: int, **fields: str) -> None:
        with self.lock:
            df = self.load()
            mask = df["id"] == int(row_id)
            if not mask.any():
                return
            for column, value in fields.items():
                df.loc[mask, column] = value
            self.save(df)

    def mark_sent(self, row_id: int) -> None:
        self.update_row(row_id, status=STATUS_SENT, sent_at=now_str(), error_message="")

    def mark_error(self, row_id: int, message: str) -> None:
        self.update_row(row_id, status=STATUS_ERROR, error_message=message)

    def get_row(self, row_id: int) -> dict | None:
        df = self.load()
        match = df[df["id"] == int(row_id)]
        return None if match.empty else match.iloc[0].to_dict()

    def next_pending(self) -> dict | None:
        df = self.load()
        pending = df[df["status"] == STATUS_PENDING]
        return None if pending.empty else pending.iloc[0].to_dict()

    def counts(self) -> dict[str, int]:
        df = self.load()
        counts = Counter(df["status"])
        return {"total": len(df), **{status: counts.get(status, 0) for status in STATUSES}}


def normalize_queue(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce any queue-shaped frame into the canonical columns and types."""
    df = df.copy()
    for column in QUEUE_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    extra = [c for c in df.columns if c not in QUEUE_COLUMNS]
    df = df[QUEUE_COLUMNS + extra].reset_index(drop=True)

    for column in df.columns:
        if column != "id":
            df[column] = df[column].fillna("").astype(str)

    ids = pd.to_numeric(df["id"], errors="coerce")
    next_id = int(ids.max()) + 1 if ids.notna().any() else 1
    for idx in ids[ids.isna()].index:
        ids.loc[idx] = next_id
        next_id += 1
    df["id"] = ids.astype(int)

    df["status"] = df["status"].str.strip().str.lower()
    df.loc[~df["status"].isin(STATUSES), "status"] = STATUS_PENDING
    df["lang"] = df["lang"].str.strip().str.lower()
    df["recipient_email"] = df["recipient_email"].str.strip()
    df["attachment_filename"] = df["attachment_filename"].str.strip()
    return df


queue_store = QueueStore(config.QUEUE_CSV)


# --------------------------------------------------------------------------- validation


def split_recipients(value: str) -> list[str]:
    return [part for part in re.split(r"[;,\s]+", value or "") if part]


def attachment_names(row: dict) -> list[str]:
    """Attachment file names for a row; falls back to the default for its language."""
    raw = (row.get("attachment_filename") or "").strip()
    if raw.lower() == NO_ATTACHMENT:
        return []
    if not raw:
        default = DEFAULT_ATTACHMENTS.get((row.get("lang") or "").lower())
        return [default] if default else []
    return [part.strip() for part in raw.split(";") if part.strip()]


def resolve_attachment(name: str) -> Path | None:
    """Path of an existing file inside attachments/, or None."""
    base = ATTACHMENTS_DIR.resolve()
    path = (base / name).resolve()
    if path.parent != base or not path.is_file():
        return None
    return path


def validate_row(row: dict) -> list[str]:
    problems = []
    recipients = split_recipients(row.get("recipient_email", ""))
    if not recipients:
        problems.append("No recipient e-mail address")
    bad = [r for r in recipients if not EMAIL_RE.match(r)]
    if bad:
        problems.append(f"Invalid e-mail address: {', '.join(bad)}")
    if not (row.get("subject") or "").strip():
        problems.append("Empty subject")
    if not (row.get("body") or "").strip():
        problems.append("Empty body")

    raw = (row.get("attachment_filename") or "").strip()
    names = attachment_names(row)
    if not names and raw.lower() != NO_ATTACHMENT:
        problems.append(
            f"No attachment set and no default for lang '{row.get('lang', '')}' "
            f"(use '{NO_ATTACHMENT}' to send without one)"
        )
    missing = [name for name in names if resolve_attachment(name) is None]
    if missing:
        problems.append(f"Attachment not found in attachments/: {', '.join(missing)}")
    return problems


# --------------------------------------------------------------------------- message building


def text_to_html(text: str) -> str:
    """Plain text -> an HTML fragment: blank lines split paragraphs, links become clickable."""

    def linkify(escaped: str) -> str:
        def repl(match: re.Match) -> str:
            if match.group("url"):
                url = match.group("url")
                return f'<a href="{url}">{url}</a>'
            address = match.group("email")
            return f'<a href="mailto:{address}">{address}</a>'

        return LINK_RE.sub(repl, escaped)

    paragraphs = [p for p in re.split(r"\n\s*\n", text.strip()) if p.strip()]
    body = "\n".join(
        '<p style="margin:0 0 12px 0">' + linkify(html.escape(p, quote=False)).replace("\n", "<br>") + "</p>"
        for p in paragraphs
    )
    return f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;color:#222">{body}</div>'


def build_message(
    row: dict,
    settings: config.SmtpSettings,
    to_override: str | None = None,
    subject_prefix: str = "",
) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"] = f"{subject_prefix}{row['subject'].strip()}"
    msg["From"] = formataddr((settings.sender_name, settings.sender_email)) if settings.sender_name else settings.sender_email
    msg["To"] = ", ".join(split_recipients(to_override or row["recipient_email"]))
    if settings.reply_to:
        msg["Reply-To"] = settings.reply_to
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=settings.sender_email.rpartition("@")[2] or None)

    body = row["body"].replace("\r\n", "\n").strip() + "\n"
    msg.set_content(body)
    msg.add_alternative(f"<!DOCTYPE html><html><body>{text_to_html(body)}</body></html>", subtype="html")

    for name in attachment_names(row):
        path = resolve_attachment(name)
        if path is None:
            raise FileNotFoundError(f"Attachment not found in attachments/: {name}")
        ctype, encoding = mimetypes.guess_type(path.name)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        msg.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)
    return msg


def send_row(row: dict, settings: config.SmtpSettings, to_override: str | None = None, subject_prefix: str = "") -> str:
    """Send one message; returns the address list it was sent to."""
    msg = build_message(row, settings, to_override, subject_prefix)
    with config.open_smtp(settings) as smtp:
        smtp.send_message(msg)
    return msg["To"]


def classify_error(exc: BaseException) -> str:
    """'row'   - this letter is at fault: mark it as error and continue.
    'halt'  - the account/config is at fault: stop the run, leave the row pending.
    'retry' - probably transient (network, 4xx): retry, then halt.
    """
    text = str(exc).lower()
    if isinstance(exc, (config.ConfigError, smtplib.SMTPAuthenticationError, smtplib.SMTPSenderRefused, smtplib.SMTPNotSupportedError)):
        return "halt"
    if isinstance(exc, (FileNotFoundError, smtplib.SMTPRecipientsRefused)):
        return "row"
    if any(word in text for word in ("quota", "limit exceeded", "sending limit", "too many", "rate limit")):
        return "halt"
    if isinstance(exc, smtplib.SMTPResponseException):
        return "retry" if 400 <= exc.smtp_code < 500 else "row"
    if isinstance(exc, OSError):  # SMTPServerDisconnected, timeouts, DNS, TLS, refused connections
        return "retry"
    return "row"


# --------------------------------------------------------------------------- dispatcher


@dataclass(frozen=True)
class SendSettings:
    interval_seconds: float = 90
    jitter_min: float = 10
    jitter_max: float = 30
    max_per_run: int = 0  # 0 = no limit

    def next_delay(self) -> float:
        low, high = sorted((self.jitter_min, self.jitter_max))
        jitter = random.uniform(low, high) * random.choice((-1, 1))
        return max(MIN_DELAY_SECONDS, self.interval_seconds + jitter)

    def delay_range(self) -> tuple[float, float]:
        low, high = sorted((self.jitter_min, self.jitter_max))
        return max(MIN_DELAY_SECONDS, self.interval_seconds - high), max(MIN_DELAY_SECONDS, self.interval_seconds + high)


class State:
    IDLE = "idle"
    SENDING = "sending"
    WAITING = "waiting"
    PAUSED = "paused"
    STOPPING = "stopping"


class Dispatcher:
    """Owns the background worker thread. One instance per Streamlit server."""

    def __init__(self, store: QueueStore):
        self.store = store
        self.settings = SendSettings()
        self.log: deque[tuple[str, str, str]] = deque(maxlen=500)
        self.smtp_status: tuple[bool, str, str] | None = None  # (ok, message, checked_at)

        self._thread: threading.Thread | None = None
        self._stop = threading.Event()
        self._pause = threading.Event()
        self._lock = threading.Lock()
        self._log_lock = threading.Lock()

        self.state = State.IDLE
        self.current: str = ""
        self.next_send_at: float = 0.0  # epoch seconds; survives Stop/Start so the delay is always honoured
        self.last_result: str = ""
        self.run_sent = 0
        self.run_errors = 0

    # ---- controls

    @property
    def is_active(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self) -> str:
        with self._lock:
            if self.is_active:
                if self._pause.is_set():
                    self._pause.clear()
                    self._log("info", "Resumed.")
                    return "Resumed."
                if self._stop.is_set():
                    return "Still stopping; try again in a moment."
                return "Already running."
            if not self.store.counts()[STATUS_PENDING]:
                return "Nothing to send: no rows with status 'pending'."
            self._stop.clear()
            self._pause.clear()
            self.run_sent = self.run_errors = 0
            self.last_result = ""
            self.state = State.SENDING
            self._thread = threading.Thread(target=self._run, name="email-dispatcher", daemon=True)
            self._thread.start()
            return "Started."

    def pause(self) -> None:
        if self.is_active and not self._pause.is_set():
            self._pause.set()
            self._log("info", "Pause requested; the current e-mail (if any) will finish first.")

    def stop(self) -> None:
        if self.is_active:
            self._stop.set()
            self._pause.clear()
            self.state = State.STOPPING
            self._log("info", "Stop requested.")

    def update_settings(self, settings: SendSettings) -> None:
        self.settings = settings

    def snapshot(self) -> dict:
        return {
            "active": self.is_active,
            "state": self.state,
            "paused": self._pause.is_set(),
            "current": self.current,
            "seconds_to_next": max(0.0, self.next_send_at - time.time()),
            "last_result": self.last_result,
            "run_sent": self.run_sent,
            "run_errors": self.run_errors,
        }

    def check_smtp(self) -> tuple[bool, str]:
        ok, message = config.test_connection()
        self.smtp_status = (ok, message, now_str())
        self._log("success" if ok else "error", f"SMTP test: {message}")
        return ok, message

    def send_test(self, row_id: int, debug_email: str) -> tuple[bool, str]:
        """Dry run: send one queue row to a debug address. The row's status is not touched."""
        row = self.store.get_row(row_id)
        if row is None:
            return False, f"Row #{row_id} not found."
        debug_email = debug_email.strip()
        if not EMAIL_RE.match(debug_email):
            return False, f"'{debug_email}' is not a valid e-mail address."
        problems = validate_row({**row, "recipient_email": debug_email})
        if problems:
            return False, "; ".join(problems)
        settings = config.load_settings()
        try:
            send_row(row, settings, to_override=debug_email, subject_prefix="[TEST] ")
        except Exception as exc:  # noqa: BLE001
            message = config.describe_smtp_error(exc, settings)
            self._log("error", f"Test e-mail for #{row_id} failed: {message}")
            return False, message
        self._log("success", f"Test e-mail for #{row_id} sent to {debug_email} ({', '.join(attachment_names(row)) or 'no attachment'}).")
        return True, f"Test e-mail for #{row_id} sent to {debug_email}."

    # ---- worker

    def _run(self) -> None:
        self._log("info", "Dispatch started.")
        result = "Finished: no pending rows left."
        try:
            while self.store.next_pending() is not None:
                if not self._wait_until(self.next_send_at):
                    result = "Stopped by user."
                    break
                row = self.store.next_pending()  # re-read: rows may have been edited while waiting
                if row is None:
                    break

                problems = validate_row(row)
                if problems:
                    message = "; ".join(problems)
                    self.store.mark_error(row["id"], message)
                    self.run_errors += 1
                    self._log("error", f"#{row['id']} skipped: {message}")
                    continue

                outcome, message = self._send_with_retry(row)
                if outcome == "halt":
                    result = message
                    break
                # The server saw this attempt, so the anti-spam delay applies even after an error.
                # It is kept across Stop/Start so restarting never skips it.
                delay = self.settings.next_delay()
                self.next_send_at = time.time() + delay
                limit = self.settings.max_per_run
                if limit and self.run_sent >= limit:
                    result = f"Stopped after the per-run limit of {limit} e-mails."
                    break
                if self.store.next_pending() is not None:
                    self._log("info", f"Next e-mail in {delay:.0f} s.")
        except Exception as exc:  # noqa: BLE001 - keep the UI informed instead of dying silently
            result = f"Worker crashed: {type(exc).__name__}: {exc}"
            self._log("error", result)
        finally:
            self.state = State.IDLE
            self.current = ""
            self.last_result = result
            self._pause.clear()
            self._log("info", f"{result} This run: {self.run_sent} sent, {self.run_errors} errors.")

    def _send_with_retry(self, row: dict) -> tuple[str, str]:
        """Returns (outcome, message) where outcome is 'sent', 'error' or 'halt'."""
        row_id, address = row["id"], row["recipient_email"]
        attempts = len(RETRY_BACKOFF_SECONDS) + 1
        for attempt in range(1, attempts + 1):
            settings = config.load_settings()
            self.state = State.SENDING
            self.current = f"#{row_id} → {address}"
            try:
                recipients = send_row(row, settings)
            except Exception as exc:  # noqa: BLE001
                message = config.describe_smtp_error(exc, settings)
                kind = classify_error(exc)
                if kind == "row":
                    self.store.mark_error(row_id, message)
                    self.run_errors += 1
                    self._log("error", f"#{row_id} → {address} failed: {message}")
                    return "error", message
                self.store.update_row(row_id, error_message=f"Last attempt {now_str()}: {message}")
                if kind == "halt" or attempt == attempts:
                    self._log("error", f"#{row_id} not sent, halting (row stays pending): {message}")
                    return "halt", f"Halted: {message}"
                wait = RETRY_BACKOFF_SECONDS[attempt - 1]
                self._log("warning", f"#{row_id} attempt {attempt} failed: {message} Retrying in {wait} s.")
                if not self._wait_until(time.time() + wait):
                    return "halt", "Stopped by user."
            else:
                self.store.mark_sent(row_id)
                self.run_sent += 1
                files = ", ".join(attachment_names(row)) or "no attachment"
                self._log("success", f"#{row_id} sent → {recipients} ({files})")
                return "sent", ""
            finally:
                self.current = ""
        return "halt", "Halted."

    def _wait_until(self, deadline: float) -> bool:
        """Sleep until `deadline`, honouring pause. Returns False if stopped."""
        while not self._stop.is_set():
            if self._pause.is_set():
                self.state = State.PAUSED
                self._stop.wait(0.5)
                continue
            remaining = deadline - time.time()
            if remaining <= 0:
                return True
            self.state = State.WAITING
            self._stop.wait(min(0.5, remaining))
        return False

    def recent_log(self, limit: int = 200) -> list[tuple[str, str, str]]:
        with self._log_lock:
            return list(self.log)[:limit]

    def _log(self, level: str, message: str) -> None:
        stamp = now_str()
        with self._log_lock:
            self.log.appendleft((stamp, level, message))
            try:
                LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
                with LOG_FILE.open("a", encoding="utf-8") as fh:
                    fh.write(f"{stamp}  {level.upper():<7}  {message}\n")
            except OSError:
                pass
