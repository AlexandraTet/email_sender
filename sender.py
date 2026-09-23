"""Queue storage, message building and the background dispatch worker."""

import html
import mimetypes
import os
import random
import re
import smtplib
import threading
import time
from collections import Counter, deque
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timedelta
from datetime import time as clock_time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.policy import SMTP as MIME_POLICY  # RFC-compliant headers (UTF-8 subjects/names), CRLF line ends
from email.utils import formataddr, formatdate, make_msgid
from html.parser import HTMLParser
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
    r"(?P<url>(?:https?://|www\.)[^\s<>\"]+[^\s<>\".,;:!?)\]»'])"
    r"|(?P<email>[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,})"
    r"|(?P<phone>\+\d{10,15})(?!\d)"
)
# Placeholders in the letter bodies, filled in when sending: the blank "«____» ____________ 2026 р."
# date line, and the address in the letter's recipient block ("E-mail: ..." under the addressee).
DATE_TOKEN = "{{date}}"
RECIPIENT_TOKEN = "{{recipient_email}}"
BLANK_ADDRESS = "____________________"
EN_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
             "September", "October", "November", "December"]
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


# One store per campaign; they are separate files, so the campaigns never touch each other.
stores = {key: QueueStore(campaign.queue_csv) for key, campaign in config.CAMPAIGNS.items()}
queue_store = stores[config.DEFAULT_CAMPAIGN]  # what parser.py writes when run from the command line


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
    if not (row.get("body") or "").strip() and not (row.get("body_html") or "").strip():
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


# --------------------------------------------------------------------------- HTML <-> plain text

DEFAULT_WRAPPER_STYLE = "font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;"
# Tags whose default browser rendering has vertical margins (a blank line in plain text).
_SPACED_TAGS = {"p", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "ul", "ol", "table"}
_BLOCK_TAGS = _SPACED_TAGS | {"div", "li", "tr", "section", "article", "header", "footer"}
_SKIP_TAGS = {"head", "title", "style", "script"}
_MARGIN_RE = re.compile(r"margin(?:-(top|bottom))?\s*:\s*([^;]+)", re.IGNORECASE)
_NUMBER_RE = re.compile(r"-?\d*\.?\d+")


def linkify(escaped: str) -> str:
    """Make bare URLs, www. addresses, e-mail addresses and +phone numbers in escaped text clickable."""

    def repl(match: re.Match) -> str:
        if match.group("url"):
            url = match.group("url")
            href = url if "://" in url else f"http://{url}"
            return f'<a href="{href}">{url}</a>'
        if match.group("email"):
            address = match.group("email")
            return f'<a href="mailto:{address}">{address}</a>'
        phone = match.group("phone")
        return f'<a href="tel:{phone}">{phone}</a>'

    return LINK_RE.sub(repl, escaped)


def letter_date(lang: str, when: datetime | None = None) -> str:
    """The date as written in the letters: "22.09.2026 р." (UA) or "22 September 2026" (EN)."""
    when = when or datetime.now()
    if lang == "en":
        return f"{when.day} {EN_MONTHS[when.month - 1]} {when.year}"
    return f"{when:%d.%m.%Y} р."


def text_to_html(text: str) -> str:
    """Plain text -> HTML fragment. Used for rows without body_html, or on request in the editor."""
    paragraphs = [p for p in re.split(r"\n\s*\n", text.strip()) if p.strip()]
    body = "\n".join(
        '<p style="margin:0 0 12px 0;">' + linkify(html.escape(p, quote=False)).replace("\n", "<br>") + "</p>"
        for p in paragraphs
    )
    return f'<div style="{DEFAULT_WRAPPER_STYLE}">\n{body}\n</div>'


def _vertical_margins(style: str, default: bool) -> tuple[bool, bool]:
    """(has top margin, has bottom margin) according to an inline style."""
    top = bottom = None
    for side, value in _MARGIN_RE.findall(style or ""):
        parts = value.split()
        if not parts:
            continue
        if not side:
            top, bottom = parts[0], parts[2] if len(parts) > 2 else parts[0]
        elif side.lower() == "top":
            top = parts[0]
        else:
            bottom = parts[0]

    def positive(value: str | None) -> bool:
        if value is None:
            return default
        number = _NUMBER_RE.match(value)
        return bool(number) and float(number.group()) > 0

    return positive(top), positive(bottom)


class _TextExtractor(HTMLParser):
    """HTML -> readable plain text.

    Blocks with a vertical margin are separated by a blank line and blocks without
    one by a single newline, which is how the Word spacing ends up in the fallback text.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.pending = 0  # newlines owed before the next text: 1 = line break, 2 = blank line
        self.fresh = True  # at the start of a line: drop leading spaces
        self.skip = 0
        self.blocks: list[tuple[str, bool]] = []  # open blocks: (tag, has bottom margin)
        self.lists: list[list] = []  # open lists: [ordered, next number]
        self.links: list[tuple[str, int]] = []
        self.cells = 0  # cells seen in the current table row
        self.in_cell = 0  # inside a table cell, line breaks become spaces: "a | b" per row
        self.cell_gap = False

    def _break(self, lines: int) -> None:
        if self.in_cell:
            self.cell_gap = True
        elif self.parts:
            self.pending = max(self.pending, lines)
        self.fresh = True

    def _emit(self, text: str) -> None:
        if self.pending:
            self.parts.append("\n" * self.pending)
            self.pending = 0
        if self.cell_gap:
            if self.parts and not self.parts[-1].endswith((" ", "\n")):
                self.parts.append(" ")
            self.cell_gap = False
        self.parts.append(text)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag in _SKIP_TAGS:
            self.skip += 1
        elif tag == "br":
            if self.in_cell:
                self._break(1)
            else:
                self._emit("\n")
                self.fresh = True
        elif tag == "a":
            self.links.append((attrs.get("href") or "", len("".join(self.parts))))
        elif tag in ("td", "th"):
            if self.cells:
                self._emit(" | ")
            self.cells += 1
            self.in_cell += 1
            self.cell_gap = False
            self.fresh = True
        elif tag in _BLOCK_TAGS:
            top, bottom = _vertical_margins(attrs.get("style", ""), default=tag in _SPACED_TAGS)
            self._break(2 if top else 1)
            self.blocks.append((tag, bottom))
            if tag == "tr":
                self.cells = 0
            elif tag in ("ul", "ol"):
                start = attrs.get("start") or "1"
                self.lists.append([tag == "ol", int(start) if start.isdigit() else 1])
            elif tag == "li":
                ordered, number = self.lists[-1] if self.lists else (False, 1)
                if self.lists:
                    self.lists[-1][1] += 1
                self._emit("   " * max(len(self.lists) - 1, 0) + (f"{number}. " if ordered else "• "))
                self.fresh = True

    def handle_endtag(self, tag):
        if tag in _SKIP_TAGS:
            self.skip = max(0, self.skip - 1)
        elif tag in ("td", "th"):
            self.in_cell = max(0, self.in_cell - 1)
            self.cell_gap = False
        elif tag == "a" and self.links:
            href, start = self.links.pop()
            label = "".join(self.parts)[start:]
            target = re.sub(r"^(mailto:|tel:|https?://)", "", href, flags=re.IGNORECASE).rstrip("/")
            if target and not href.startswith("#") and target not in label:
                self._emit(f" ({href.removeprefix('mailto:').removeprefix('tel:')})")
        elif tag in _BLOCK_TAGS:
            for i in range(len(self.blocks) - 1, -1, -1):
                if self.blocks[i][0] == tag:
                    closed = self.blocks[i:]
                    del self.blocks[i:]
                    for name, _ in closed:
                        if name in ("ul", "ol") and self.lists:
                            self.lists.pop()
                    self._break(2 if closed[0][1] else 1)
                    break

    def handle_data(self, data):
        if self.skip:
            return
        text = re.sub(r"[ \t\r\n\f]+", " ", data)  # HTML whitespace; &nbsp; (\xa0) is kept
        if self.fresh:
            text = text.lstrip(" ")
            if not text.strip():
                return
        if text:
            self._emit(text)
            self.fresh = False

    def text(self) -> str:
        out = "".join(self.parts).replace("\xa0", " ").replace(" ", "\t")
        out = "\n".join(line.rstrip() for line in out.split("\n"))
        return re.sub(r"\n{3,}", "\n\n", out).strip()


def html_to_text(markup: str) -> str:
    extractor = _TextExtractor()
    extractor.feed(markup or "")
    extractor.close()
    return extractor.text()


def stored_bodies(row: dict) -> tuple[str, str]:
    """(plain text, HTML fragment) as stored, with placeholders; a missing one is derived from the other."""
    text = (row.get("body") or "").replace("\r\n", "\n").strip()
    fragment = (row.get("body_html") or "").strip()
    if not fragment and text:
        fragment = text_to_html(text)
    if not text and fragment:
        text = html_to_text(fragment)
    return text, fragment


def message_bodies(row: dict, when: datetime | None = None) -> tuple[str, str]:
    """(plain text, HTML fragment) exactly as sent: {{date}} and {{recipient_email}} filled in."""
    date = letter_date(row.get("lang", ""), when)
    addresses = split_recipients(row.get("recipient_email", ""))
    links = ", ".join(f'<a href="mailto:{html.escape(a)}">{html.escape(a)}</a>' for a in addresses)
    text, fragment = stored_bodies(row)
    text = text.replace(DATE_TOKEN, date).replace(RECIPIENT_TOKEN, ", ".join(addresses) or BLANK_ADDRESS)
    fragment = fragment.replace(DATE_TOKEN, html.escape(date)).replace(RECIPIENT_TOKEN, links or BLANK_ADDRESS)
    return text, fragment


def html_document(fragment: str, title: str = "", lang: str = "") -> str:
    """A complete, e-mail-safe HTML document: inline styles only, no <style> blocks or classes."""
    lang_attr = {"ua": "uk", "en": "en"}.get(lang, "")
    return (
        "<!DOCTYPE html>\n"
        + (f'<html lang="{lang_attr}">\n' if lang_attr else "<html>\n")
        + "<head>\n"
        '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        '<meta name="x-apple-disable-message-reformatting">\n'
        f"<title>{html.escape(title)}</title>\n"
        "</head>\n"
        '<body style="margin:0;padding:0;">\n'
        f"{fragment}\n"
        "</body>\n</html>\n"
    )


# --------------------------------------------------------------------------- message building


def build_message(
    row: dict,
    settings: config.SmtpSettings,
    to_override: str | None = None,
    subject_prefix: str = "",
) -> MIMEMultipart:
    """multipart/alternative (plain + HTML), wrapped in multipart/mixed when there are attachments."""
    text, fragment = message_bodies(row)
    subject = " ".join(f"{subject_prefix}{row['subject']}".split())

    body = MIMEMultipart("alternative", policy=MIME_POLICY)
    # Plain text first: clients show the last alternative they can render.
    body.attach(MIMEText(text + "\n", "plain", "utf-8", policy=MIME_POLICY))
    body.attach(MIMEText(html_document(fragment, subject, row.get("lang", "")), "html", "utf-8", policy=MIME_POLICY))

    files = []
    for name in attachment_names(row):
        path = resolve_attachment(name)
        if path is None:
            raise FileNotFoundError(f"Attachment not found in attachments/: {name}")
        files.append(path)

    if files:
        msg = MIMEMultipart("mixed", policy=MIME_POLICY)
        msg.attach(body)
        for path in files:
            ctype, encoding = mimetypes.guess_type(path.name)
            if ctype is None or encoding is not None:
                ctype = "application/octet-stream"
            maintype, subtype = ctype.split("/", 1)
            part = MIMEBase(maintype, subtype, policy=MIME_POLICY)
            part.set_payload(path.read_bytes())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=path.name)
            msg.attach(part)
    else:
        msg = body

    msg["Subject"] = subject
    msg["From"] = formataddr((settings.sender_name, settings.sender_email)) if settings.sender_name else settings.sender_email
    msg["To"] = ", ".join(split_recipients(to_override or row["recipient_email"]))
    if settings.reply_to:
        msg["Reply-To"] = settings.reply_to
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=settings.sender_email.rpartition("@")[2] or None)
    return msg


def send_row(row: dict, settings: config.SmtpSettings, to_override: str | None = None, subject_prefix: str = "") -> str:
    """Send one message; returns the address list it was sent to."""
    msg = build_message(row, settings, to_override, subject_prefix)
    with config.open_smtp(settings) as smtp:
        smtp.send_message(msg)
    return str(msg["To"])


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


def active_campaign() -> str:
    """Name of the campaign sending right now, or "" - only one may send at a time."""
    running = Dispatcher._active
    return running.name if running is not None and running.is_active else ""


def next_occurrence(at: clock_time, now: datetime | None = None) -> datetime:
    """The next time the clock shows `at`: today if that is still ahead, otherwise tomorrow.

    Setting 10:00 at 22:00 therefore means tomorrow morning, never "right now".
    """
    now = now or datetime.now()
    target = datetime.combine(now.date(), at)
    return target if target > now else target + timedelta(days=1)


class State:
    IDLE = "idle"
    SCHEDULED = "scheduled"  # counting down to a scheduled start
    SENDING = "sending"
    WAITING = "waiting"  # between e-mails
    PAUSED = "paused"
    STOPPING = "stopping"


class Dispatcher:
    """Owns the background worker thread. One instance per campaign per Streamlit server."""

    _active: "Dispatcher | None" = None  # the campaign currently sending; only one at a time

    def __init__(self, store: QueueStore, name: str = ""):
        self.store = store
        self.name = name
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
        self.scheduled_for: float = 0.0  # epoch seconds of a pending scheduled start, 0 = none
        self.last_result: str = ""
        self.run_sent = 0
        self.run_errors = 0

    # ---- controls

    @property
    def is_active(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self, start_at: datetime | None = None) -> str:
        """Start a run now, or at `start_at` (local time) if that is in the future. Resumes if paused."""
        with self._lock:
            if self.is_active:
                if self._pause.is_set():
                    self._pause.clear()
                    self._log("info", "Resumed.")
                    return "Resumed."
                if self._stop.is_set():
                    return "Still stopping; try again in a moment."
                return "Already running."
            busy = Dispatcher._active
            if busy is not None and busy is not self and busy.is_active:
                return f"'{busy.name or 'another campaign'}' is still sending; stop it first."
            if not self.store.counts()[STATUS_PENDING]:
                return "Nothing to send: no rows with status 'pending'."
            self._stop.clear()
            self._pause.clear()
            self.run_sent = self.run_errors = 0
            self.last_result = ""
            scheduled = start_at is not None and start_at.timestamp() > time.time()
            self.scheduled_for = start_at.timestamp() if scheduled else 0.0
            self.state = State.SCHEDULED if scheduled else State.SENDING
            Dispatcher._active = self
            self._thread = threading.Thread(target=self._run, name=f"dispatcher-{self.name or 'queue'}", daemon=True)
            self._thread.start()
            return f"Scheduled: sending starts at {start_at:%H:%M} on {start_at:%d.%m.%Y}." if scheduled else "Started."

    def start_now(self) -> None:
        """Skip the rest of a scheduled start's countdown."""
        if self.is_active and self.scheduled_for:
            self.scheduled_for = time.time()
            self._pause.clear()
            self._log("info", "Starting now instead of at the scheduled time.")

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
            "scheduled_for": self.scheduled_for,
            "seconds_to_start": max(0.0, self.scheduled_for - time.time()) if self.scheduled_for else 0.0,
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
        result = "Finished: no pending rows left."
        try:
            if self.scheduled_for:
                self._log("info", f"Scheduled start: sending begins at {datetime.fromtimestamp(self.scheduled_for):%d.%m.%Y %H:%M}.")
                # Re-read the deadline every tick so "Start now" can move it.
                if not self._wait_until(lambda: self.scheduled_for, State.SCHEDULED):
                    result = "Scheduled start cancelled."
                    return
                self.scheduled_for = 0.0
            self._log("info", "Dispatch started.")
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
            if Dispatcher._active is self:
                Dispatcher._active = None
            self.state = State.IDLE
            self.current = ""
            self.scheduled_for = 0.0
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

    def _wait_until(self, deadline: float | Callable[[], float], state: str = State.WAITING) -> bool:
        """Sleep until `deadline` (epoch seconds, or a function returning it), honouring pause.

        While paused the clock keeps running, but nothing proceeds until Resume, even if the
        deadline has passed in the meantime. Returns False if stopped.
        """
        get_deadline = deadline if callable(deadline) else (lambda: deadline)
        while not self._stop.is_set():
            if self._pause.is_set():
                self.state = State.PAUSED
                self._stop.wait(0.5)
                continue
            remaining = get_deadline() - time.time()
            if remaining <= 0:
                return True
            self.state = state
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
                with LOG_FILE.open("a", encoding="utf-8") as fh:  # shared by all campaigns
                    fh.write(f"{stamp}  {level.upper():<7}  {f'[{self.name}] ' if self.name else ''}{message}\n")
            except OSError:
                pass
