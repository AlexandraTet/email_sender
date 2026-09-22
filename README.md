# Email Dispatcher

Sends the letters in `emails.docx` one by one over SMTP, each with its PDF from
`attachments/`, at a randomised, spam-friendly pace. A Streamlit dashboard lets
you edit the queue, send a test first, start/pause/stop, and watch progress live.

## Quick start (Windows)

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.template .env
streamlit run app.py
```

Fill in `.env` before starting (see [App passwords](#app-passwords)). The dashboard opens at
http://localhost:8501. Tested with Python 3.14, Streamlit 1.64 and pandas 3.0.

## Workflow

1. **Import.** On first launch `emails.docx` is parsed into `data/queue.csv`. The current file
   gives 252 letters (117 UA, 135 EN). 35 letters have no address in the document
   (`E-mail: ____`); they are marked `error`. Fill in the address in the **Queue** or
   **Letter editor** tab and set the status to `pending`.
2. **Check.** Click **Test SMTP**, then open the **Pre-send check** tab: it lists pending rows
   that would fail, missing attachments and addresses used by several letters.
3. **Dry run.** In the sidebar, send any letter to your own address. It is sent with the real
   attachment and a `[TEST]` subject prefix, and its queue status does not change.
4. **Send.** Set the interval and jitter, then **Start** and confirm.
   - **Pause** finishes the current e-mail, then waits. While paused you can edit the queue, and edits are used when you **Resume**.
   - **Stop** ends the run.
   - Only `pending` rows are sent.
5. **Monitor.** The progress bar, counters, countdown to the next e-mail and the activity
   log update every second. The log is also written to `data/activity.log`.

To re-import after editing the `.docx`, use **Re-import emails.docx**. Letters already sent keep
their `sent` status (matched by recipient + subject), so nothing is sent twice. The old
queue is saved to `data/queue.backup.csv`. From the command line: `python parser.py` (or
`python parser.py --fresh` to reset everything to pending).

## How it works

| Topic | Behaviour |
|---|---|
| Parsing | Each letter starts at the university letterhead line, which also gives the language (UA/EN). Then come the addressee, `E-mail: …` → `recipient_email`, the next line → `subject`, and the salutation to the end → `body`. Organization names come from the table of contents. |
| Attachments | Empty = default for the language (`expense_estimate_ua.pdf` / `expense_estimate_en.pdf`, set in `config.py`). Several files: `a.pdf; b.xlsx`. No attachment: `none`. A letter can name its own attachment with a line such as `Attachment: file.pdf`. Any file type works (the MIME type is guessed). |
| Validation | Before each send the recipient, subject, body and every attachment are checked. A failing row becomes `error` with the reason in `error_message`. |
| Pace | Delay = interval ± random jitter (default 90 s ± 10–30 s, minimum 5 s). The delay is honoured across Stop/Start. **Stop after N e-mails** helps you stay under daily limits. |
| Errors | Problems with a single letter (bad address, rejected recipient, missing file) mark that row `error` and the run continues. Account problems (wrong password, quota exceeded) halt the run and leave the row `pending`. Network errors are retried twice (after 30 s and 90 s), then the run halts. |
| Format | Each e-mail has a plain-text part plus an HTML version with clickable links, and proper `Date`/`Message-ID` headers. `SENDER_NAME` and `REPLY_TO` are optional. |
| Queue file | `data/queue.csv` is UTF-8 with a BOM, so it opens in Excel. Close Excel before the app writes to it. |

## App passwords

These providers reject your normal password for SMTP. Create an app password and put it in
`SENDER_PASSWORD`.

**Gmail** (`smtp.gmail.com`, port 465 SSL or 587 STARTTLS)
1. Turn on 2-Step Verification: Google Account → Security.
2. Open <https://myaccount.google.com/apppasswords>, type a name (e.g. *Email Dispatcher*), click **Create**.
3. Copy the 16-character password; spaces don't matter.

Personal Gmail allows roughly 500 messages per day. Messages sent this way also appear in your Sent folder.

**Ukr.net** (`smtp.ukr.net`, port 465 SSL)
1. Sign in at mail.ukr.net → **Налаштування** (Settings) → the mail-programs / IMAP access section.
2. Enable access for external programs (IMAP/SMTP) and generate a password for the app.
3. Use that generated password. Menu names may differ slightly between interface versions.

**Outlook / Hotmail** (`smtp-mail.outlook.com`, port 587 STARTTLS)
1. At <https://account.microsoft.com/security>, turn on two-step verification (Advanced security options).
2. Under **App passwords**, create a new one and use it.

Microsoft is retiring password-based SMTP for Outlook.com. If **Test SMTP** reports an
authentication error even with a correct app password, the account needs OAuth, which this
tool does not support. Use Gmail or Ukr.net instead.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit dashboard |
| `sender.py` | Queue storage, validation, message building, background worker |
| `parser.py` | `emails.docx` → `data/queue.csv` |
| `config.py` | Paths, `.env` loading, SMTP connection (SSL/STARTTLS) and error messages |
| `.env.template` | Settings template; copy it to `.env` (git-ignored) |

## Troubleshooting

- **Timeout / "Server closed the connection" / TLS error.** The port and security mode don't match. Use 465 with SSL or 587 with STARTTLS; `SMTP_SECURITY=auto` picks the right one for the port.
- **"Cannot write queue.csv".** The file is open in Excel. Close it.
- **Closing the terminal (Ctrl+C) stops sending.** Rows already sent keep their status; run the app again and press Start to continue.
- **Run one copy of the app at a time.** Two copies would both send the same queue.
- `.env` changes apply without a restart. Click **Test SMTP** again after editing it.
