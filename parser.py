"""Parse emails.docx into data/queue.csv.

Expected layout of emails.docx (what the current file uses):

    Table of contents:  "1. Sigma Software Labs  [ENG]", "2. Intellias  [ENG]", ...
    Then, per letter:
        letterhead            "NATIONAL TECHNICAL UNIVERSITY OF UKRAINE" (EN) or
                              "НАЦІОНАЛЬНИЙ ТЕХНІЧНИЙ УНІВЕРСИТЕТ УКРАЇНИ" (UA), then address lines
        date line             "«____» ____________ 2026"
        addressee (+ phone)   "To the Management of Sigma Software Labs"
        e-mail line           "E-mail: info@sigma.software"     -> recipient_email
        subject line          "Request for sponsorship ..."     -> subject
        salutation ... end    "Dear Sigma Software Labs Team," ... contacts   -> body

An explicit attachment can be named anywhere in a letter with a line such as
"Attachment: brochure_en.pdf" (or "Вкладення:", "Додаток:", "Файл:"). Otherwise the
attachment defaults to config.DEFAULT_ATTACHMENTS[lang].

Run standalone:  python parser.py            (re-import, keep 'sent' history)
                 python parser.py --fresh    (re-import, everything back to pending)
"""

from __future__ import annotations

import argparse
import re
import shutil
from collections import defaultdict, deque
from pathlib import Path

import docx
import pandas as pd
from docx.text.paragraph import Paragraph

import config
from config import (
    ATTACHMENTS_DIR,
    DEFAULT_ATTACHMENTS,
    QUEUE_BACKUP_CSV,
    QUEUE_COLUMNS,
    STATUS_ERROR,
    STATUS_PENDING,
    STATUS_SENT,
)
from sender import queue_store

LETTERHEADS = {
    "NATIONAL TECHNICAL UNIVERSITY OF UKRAINE": "en",
    "НАЦІОНАЛЬНИЙ ТЕХНІЧНИЙ УНІВЕРСИТЕТ УКРАЇНИ": "ua",
}
# "E-mail:", "Email:", "e-mail :" - also with a Cyrillic "Е".
EMAIL_LINE_RE = re.compile(r"^\s*[EЕeе]\s*[-‐‑–]?\s*mail\s*:\s*(.*)$", re.IGNORECASE)
ADDRESS_RE = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")
ATTACHMENT_LINE_RE = re.compile(r"^\s*(?:attachments?|attached|вкладення|додаток|додатки|файл)\s*:\s*(.+)$", re.IGNORECASE)
TOC_RE = re.compile(r"^\s*(\d+)\.\s+(.+?)\s*(?:\[(ENG|EN|UKR|UA)\])?\s*$", re.IGNORECASE)
DATE_LINE_RE = re.compile(r"(?:«.*»|\d{1,2}\.\d{1,2}\.)\s*.*20\d\d")
PHONE_LINE_RE = re.compile(r"^\s*(?:тел|tel|phone)", re.IGNORECASE)
CYRILLIC_RE = re.compile(r"[а-яіїєґ]", re.IGNORECASE)

MISSING_EMAIL_MESSAGE = "No e-mail address in emails.docx: fill in recipient_email and set status to pending"
ADDRESSEE_PREFIXES = re.compile(
    r"^(to the management of|to the|керівництву компанії|керівництву|директору)\s+", re.IGNORECASE
)


def _space(value) -> int:
    return int(value or 0)


def paragraphs_to_text(paragraphs: list[Paragraph]) -> str:
    """Join paragraphs, using the document's own spacing to decide where blank lines go.

    Paragraphs followed by vertical space (every prose paragraph, the salutation,
    the end of each signature block) end with a blank line; tightly stacked lines
    such as signatures and contact details stay on consecutive lines.
    """
    texts = [p.text.strip() for p in paragraphs]
    out: list[str] = []
    for i, (para, text) in enumerate(zip(paragraphs, texts)):
        if not text:
            continue
        out.append(text)
        nxt = paragraphs[i + 1] if i + 1 < len(paragraphs) else None
        gap = _space(para.paragraph_format.space_after) > 0 or (
            nxt is not None and _space(nxt.paragraph_format.space_before) > 0
        )
        out.append("\n\n" if gap else "\n")
    return "".join(out).strip()


def _detect_lang(letterhead_lang: str | None, toc_marker: str | None, sample: str) -> str:
    if letterhead_lang:
        return letterhead_lang
    if toc_marker:
        return "en" if toc_marker.lower().startswith("en") else "ua"
    letters = [c for c in sample if c.isalpha()]
    cyrillic = sum(1 for c in letters if CYRILLIC_RE.match(c))
    return "ua" if letters and cyrillic / len(letters) > 0.3 else "en"


def _find_attachment(texts: list[str]) -> tuple[str, set[int]]:
    """Return an explicitly named attachment (';'-joined) and the indexes of its lines."""
    found: list[str] = []
    lines: set[int] = set()
    existing = {p.name.lower(): p.name for p in ATTACHMENTS_DIR.glob("*") if p.is_file()}
    for i, text in enumerate(texts):
        match = ATTACHMENT_LINE_RE.match(text)
        if not match:
            continue
        lines.add(i)
        for name in re.split(r"[;,]", match.group(1)):
            name = name.strip().strip("«»\"'")
            if name:
                found.append(existing.get(name.lower(), name))
    return "; ".join(dict.fromkeys(found)), lines


def parse_docx(path: Path = config.DOCX_PATH) -> pd.DataFrame:
    document = docx.Document(str(path))
    paragraphs = document.paragraphs
    texts = [p.text.strip() for p in paragraphs]

    starts = [i for i, t in enumerate(texts) if t in LETTERHEADS]
    if not starts:
        raise ValueError(
            "No letters found: expected each letter to start with the letterhead line "
            f"{' or '.join(repr(k) for k in LETTERHEADS)}."
        )

    toc = [TOC_RE.match(t) for t in texts[: starts[0]]]
    toc = [m for m in toc if m]

    rows = []
    bounds = starts + [len(paragraphs)]
    for n, start in enumerate(starts):
        end = bounds[n + 1]
        block = list(range(start, end))
        email_idx = next((i for i in block[:25] if EMAIL_LINE_RE.match(texts[i])), None)
        if email_idx is None:
            raise ValueError(f"Letter {n + 1} (paragraph {start}) has no 'E-mail:' line near its top.")

        addresses = ADDRESS_RE.findall(EMAIL_LINE_RE.match(texts[email_idx]).group(1))
        date_idx = max((i for i in range(start, email_idx) if DATE_LINE_RE.search(texts[i])), default=start)
        addressee = [texts[i] for i in range(date_idx + 1, email_idx) if texts[i] and not PHONE_LINE_RE.match(texts[i])]

        rest = [i for i in range(email_idx + 1, end) if texts[i]]
        if not rest:
            raise ValueError(f"Letter {n + 1} (paragraph {start}) has no subject/body after the e-mail line.")
        subject_idx, body_idx = rest[0], rest[1:]

        attachment, attachment_lines = _find_attachment([texts[i] for i in body_idx])
        body_paragraphs = [paragraphs[i] for k, i in enumerate(body_idx) if k not in attachment_lines]
        body = paragraphs_to_text(body_paragraphs)

        toc_match = toc[n] if len(toc) == len(starts) else None
        lang = _detect_lang(LETTERHEADS.get(texts[start]), toc_match and toc_match.group(3), texts[subject_idx] + body[:500])
        if toc_match:
            organization = toc_match.group(2).strip()
        else:
            organization = ADDRESSEE_PREFIXES.sub("", addressee[0]).strip("«» ") if addressee else ""

        rows.append(
            {
                "id": n + 1,
                "organization": organization,
                "recipient_email": ", ".join(dict.fromkeys(addresses)),
                "subject": " ".join(texts[subject_idx].split()),
                "body": body,
                "lang": lang,
                "attachment_filename": attachment or DEFAULT_ATTACHMENTS.get(lang, ""),
                "status": STATUS_PENDING if addresses else STATUS_ERROR,
                "error_message": "" if addresses else MISSING_EMAIL_MESSAGE,
                "sent_at": "",
            }
        )
    return pd.DataFrame(rows, columns=QUEUE_COLUMNS)


def _history_key(row) -> tuple[str, str]:
    return (
        " ".join(str(row["recipient_email"]).lower().split()),
        " ".join(str(row["subject"]).lower().split()),
    )


def carry_over_sent(new: pd.DataFrame, old: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Keep 'sent' status for letters that were already sent (matched by recipient + subject).

    Matching on content rather than id means reordering letters in the .docx never
    causes a second e-mail to someone who already got one.
    """
    history: dict[tuple[str, str], deque] = defaultdict(deque)
    for _, row in old[old["status"] == STATUS_SENT].iterrows():
        history[_history_key(row)].append(row)

    new = new.copy()
    kept = 0
    for idx, row in new.iterrows():
        queue = history.get(_history_key(row))
        if row["recipient_email"] and queue:
            previous = queue.popleft()
            new.loc[idx, ["status", "sent_at", "error_message"]] = [STATUS_SENT, previous["sent_at"], ""]
            kept += 1
    return new, kept


def import_docx(docx_path: Path = config.DOCX_PATH, keep_sent: bool = True) -> dict:
    """Parse the .docx and (re)write the queue. Returns a summary for the UI."""
    new = parse_docx(docx_path)
    kept = 0
    if queue_store.exists():
        shutil.copyfile(queue_store.path, QUEUE_BACKUP_CSV)
        if keep_sent:
            new, kept = carry_over_sent(new, queue_store.load())
    queue_store.save(new)
    return {
        "total": len(new),
        "ua": int((new["lang"] == "ua").sum()),
        "en": int((new["lang"] == "en").sum()),
        "missing_email": int((new["recipient_email"] == "").sum()),
        "kept_sent": kept,
        "pending": int((new["status"] == STATUS_PENDING).sum()),
    }


def main() -> None:
    cli = argparse.ArgumentParser(description="Parse emails.docx into data/queue.csv")
    cli.add_argument("--docx", type=Path, default=config.DOCX_PATH)
    cli.add_argument("--fresh", action="store_true", help="do not keep 'sent' status from the existing queue")
    args = cli.parse_args()
    summary = import_docx(args.docx, keep_sent=not args.fresh)
    print(
        f"{summary['total']} letters (UA {summary['ua']}, EN {summary['en']}) -> {config.QUEUE_CSV}\n"
        f"pending: {summary['pending']}, missing e-mail: {summary['missing_email']}, "
        f"kept as sent: {summary['kept_sent']}"
    )


if __name__ == "__main__":
    main()
