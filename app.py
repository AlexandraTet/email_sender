"""Streamlit dashboard for the e-mail dispatcher.  Run:  streamlit run app.py"""

from __future__ import annotations

import importlib
import os
import sys
import threading
import time
from datetime import datetime, timedelta
from datetime import time as clock_time

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Email Dispatcher", page_icon="✉️", layout="wide")

LOCAL_MODULES = ("config", "sender", "parser")  # dependency order
ATTEMPTS = 3


@st.cache_resource(show_spinner=False)
def _import_lock() -> threading.Lock:
    """One lock for the whole server: every browser session runs its own script thread."""
    return threading.Lock()


def _load_local_modules(force: bool = False) -> None:
    """Import config/sender/parser, dropping copies left behind by an edit.

    Streamlit only reloads local modules it already watches, and refreshes that list only after a
    successful run, so editing a file while the app runs can leave an outdated module in
    sys.modules ("cannot import name 'html_to_text' from 'sender'") until the server restarts.
    Each module is stamped with its file's mtime below; if any file changed since, all three are
    dropped, because they import from each other.

    Dropping a module while another thread imports it breaks that import half-way, so every local
    import below runs under _import_lock(). Streamlit evicts modules from other sessions' threads,
    outside that lock, so the whole block is retried if that spoils an import.
    """
    stale = any(
        module is not None and getattr(module, "__loaded_mtime__", None) != os.path.getmtime(module.__file__)
        for module in (sys.modules.get(name) for name in LOCAL_MODULES)
    )
    if stale or force:
        for name in LOCAL_MODULES:
            sys.modules.pop(name, None)
    for name in LOCAL_MODULES:
        module = importlib.import_module(name)
        module.__dict__.setdefault("__loaded_mtime__", os.path.getmtime(module.__file__))


for _attempt in range(1, ATTEMPTS + 1):
    try:
        with _import_lock():  # no other session may reload these modules while we import them
            _load_local_modules(force=_attempt > 1)
            import config  # noqa: E402
            import parser as docx_parser  # noqa: E402
            from config import (  # noqa: E402
                ATTACHMENTS_DIR,
                CAMPAIGNS,
                DEFAULT_CAMPAIGN,
                LANGS,
                NO_ATTACHMENT,
                QUEUE_COLUMNS,
                STATUS_ERROR,
                STATUS_PENDING,
                STATUS_SENT,
                STATUSES,
            )
            from sender import (  # noqa: E402
                DATE_TOKEN,
                RECIPIENT_TOKEN,
                Dispatcher,
                QueueLockedError,
                SendSettings,
                State,
                active_campaign,
                attachment_names,
                html_to_text,
                message_bodies,
                next_occurrence,
                now_str,
                resolve_attachment,
                split_recipients,
                stored_bodies,
                stores,
                text_to_html,
                validate_row,
            )
        break
    except (ImportError, AttributeError, KeyError) as _exc:  # another session reloaded mid-import
        if _attempt == ATTEMPTS:
            st.error(
                f"Could not load the app's own modules ({type(_exc).__name__}: {_exc}). This happens when "
                "config.py / sender.py / parser.py change while the app is running: press R to rerun, "
                "or restart `streamlit run app.py`."
            )
            st.stop()
        time.sleep(0.05)

# Letter bodies are edited in the Letter editor tab, where HTML and plain text stay in sync.
EDITABLE_COLUMNS = ["organization", "recipient_email", "subject", "lang", "attachment_filename", "status"]
BODY_VIEWS = ["👁️ HTML view", "</> HTML source", "📝 Plain text"]
EDITOR_COLUMN_ORDER = [
    "id", "status", "organization", "recipient_email", "subject", "lang",
    "attachment_filename", "body", "error_message", "sent_at",
]
LOG_ICONS = {"info": "ℹ️", "success": "✅", "warning": "⚠️", "error": "❌"}
STATUS_ICONS = {STATUS_PENDING: "🕓", STATUS_SENT: "✅", STATUS_ERROR: "❌"}
SCHEDULE_MODES = ["At a time", "After a delay"]
DEFAULT_START_TIME = clock_time(10, 0)


@st.cache_resource
def _server_state() -> dict:
    """Server-wide state that survives reruns and module reloads."""
    return {}


def get_dispatcher(campaign) -> Dispatcher:
    """One worker per campaign per Streamlit server, shared by all browser tabs and reruns.

    After sender.py is reloaded the worker is rebuilt from the new code, but never while a run is
    in progress, so there is only ever one sending thread per campaign (and sender only lets one
    campaign send at a time).
    """
    state = _server_state().setdefault("dispatchers", {})
    current = state.get(campaign.key)
    if current is None or (not isinstance(current, Dispatcher) and not current.is_active):
        fresh = Dispatcher(stores[campaign.key], name=campaign.key)
        if current is not None:
            fresh.log.extend(current.log)
            fresh.smtp_status, fresh.last_result = current.smtp_status, current.last_result
            fresh.next_send_at = current.next_send_at  # keep the anti-spam delay
        state[campaign.key] = fresh
    return state[campaign.key]


ss = st.session_state
ss.setdefault("editor_version", 0)
ss.setdefault("letter_version", 0)


# --------------------------------------------------------------------------- helpers


def fmt_duration(seconds: float) -> str:
    seconds = int(round(seconds))
    hours, rest = divmod(seconds, 3600)
    minutes, secs = divmod(rest, 60)
    if hours:
        return f"{hours}h {minutes:02d}m"
    return f"{minutes}m {secs:02d}s" if minutes else f"{secs}s"


def _campaign_summary(key: str) -> str:
    counts = stores[key].counts()
    return f"{counts[STATUS_SENT]}/{counts['total']} sent" if counts["total"] else "not imported"


def fmt_countdown(seconds: float) -> str:
    hours, rest = divmod(max(0, int(seconds + 0.999)), 3600)  # round up: never shows 00s before the start
    minutes, secs = divmod(rest, 60)
    return f"{hours:02d}h {minutes:02d}m {secs:02d}s"


def fmt_when(moment: datetime) -> str:
    """'10:00', '10:00 tomorrow' or '10:00 on 25.09'."""
    days = (moment.date() - datetime.now().date()).days
    return f"{moment:%H:%M}" + ("" if days == 0 else " tomorrow" if days == 1 else f" on {moment:%d.%m}")


def scheduled_start(now: datetime | None = None) -> datetime | None:
    """When a run started now should begin, per the sidebar; None = immediately."""
    if not ss.get("sched_on"):
        return None
    now = now or datetime.now()
    if ss.get("sched_mode") == SCHEDULE_MODES[1]:
        amount = ss.get("sched_delay", 60)
        return now + (timedelta(hours=amount) if ss.get("sched_unit") == "hours" else timedelta(minutes=amount))
    return next_occurrence(ss.get("sched_time", DEFAULT_START_TIME), now)


def flash(kind: str, message: str) -> None:
    """Show a message on the next run (survives st.rerun)."""
    ss["flash"] = (kind, message)


def is_locked(snap: dict) -> bool:
    """The queue is read-only while the worker may write to it (not idle, paused or counting down)."""
    return snap["active"] and snap["state"] not in (State.PAUSED, State.SCHEDULED)


def attachment_files() -> list[str]:
    if not ATTACHMENTS_DIR.exists():
        return []
    return sorted(p.name for p in ATTACHMENTS_DIR.iterdir() if p.is_file() and not p.name.startswith("."))


def row_label(df: pd.DataFrame):
    labels = {
        int(r.id): f"#{r.id} · {STATUS_ICONS.get(r.status, '')} {r.organization[:45] or r.recipient_email or '—'} · {r.lang.upper()}"
        for r in df.itertuples()
    }
    return lambda row_id: labels.get(int(row_id), f"#{row_id}")


def editor_key() -> str:
    # Campaign and status filter are part of the key, so saved edits always refer to the rows shown.
    shown = ss.get("status_filter") or STATUSES
    return f"queue_editor_{ss.get('campaign', DEFAULT_CAMPAIGN)}_{ss.editor_version}_{'-'.join(sorted(shown))}"


def pending_edits() -> dict | None:
    state = ss.get(editor_key())
    if state and (state.get("edited_rows") or state.get("added_rows") or state.get("deleted_rows")):
        return state
    return None


def reset_editor() -> None:
    ss.pop("editor_base", None)
    ss.editor_version += 1


def apply_editor_changes(view: pd.DataFrame, edits: dict) -> int:
    """Apply only the cells the user changed onto the *current* file (3-way merge by id).

    The worker may have updated statuses since the editor was loaded; those rows
    are left alone unless the user edited the same cell.
    """
    changes = 0
    with store.lock:
        df = store.load()
        for position, cells in edits.get("edited_rows", {}).items():
            row_id = int(view.iloc[int(position)]["id"])
            mask = df["id"] == row_id
            if not mask.any():
                continue
            for column, value in cells.items():
                if column not in EDITABLE_COLUMNS:
                    continue
                value = "" if value is None else str(value)
                if column != "body":
                    value = value.strip()
                if column == "lang" and "attachment_filename" not in cells:
                    # Follow the language switch if the row still uses the default attachment.
                    old_lang = df.loc[mask, "lang"].iloc[0]
                    if df.loc[mask, "attachment_filename"].iloc[0] == campaign.attachments.get(old_lang):
                        df.loc[mask, "attachment_filename"] = campaign.attachments.get(value, "")
                if column == "status":
                    if value == STATUS_PENDING:
                        df.loc[mask, ["error_message", "sent_at"]] = ""
                    elif value == STATUS_SENT and not df.loc[mask, "sent_at"].iloc[0]:
                        df.loc[mask, "sent_at"] = f"{now_str()} (marked manually)"
                df.loc[mask, column] = value
                changes += 1

        deleted = [int(view.iloc[int(p)]["id"]) for p in edits.get("deleted_rows", [])]
        if deleted:
            df = df[~df["id"].isin(deleted)]
            changes += len(deleted)

        next_id = int(df["id"].max()) + 1 if len(df) else 1
        new_rows = []
        for added in edits.get("added_rows", []):
            row = {column: "" for column in QUEUE_COLUMNS}
            row.update({k: "" if v is None else str(v) for k, v in added.items() if k in EDITABLE_COLUMNS})
            row["id"] = next_id
            row["lang"] = row["lang"] if row["lang"] in LANGS else "en"
            row["status"] = row["status"] if row["status"] in STATUSES else STATUS_PENDING
            row["attachment_filename"] = row["attachment_filename"] or campaign.attachments.get(row["lang"], "")
            new_rows.append(row)
            next_id += 1
        if new_rows:
            df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
            changes += len(new_rows)

        store.save(df)
    return changes


def save_editor_changes() -> None:
    edits = pending_edits()
    if not edits:
        return
    try:
        changes = apply_editor_changes(ss.editor_view, edits)
    except QueueLockedError as exc:
        flash("error", str(exc))
        return
    flash("success", f"Saved {changes} change(s) to data/queue.csv.")
    reset_editor()


def reset_errors() -> None:
    with store.lock:
        df = store.load()
        mask = (df["status"] == STATUS_ERROR) & (df["recipient_email"] != "")
        df.loc[mask, ["status", "error_message"]] = [STATUS_PENDING, ""]
        skipped = int(((df["status"] == STATUS_ERROR) & (df["recipient_email"] == "")).sum())
        try:
            store.save(df)
        except QueueLockedError as exc:
            flash("error", str(exc))
            return
    note = f" {skipped} row(s) without an e-mail address stay in error." if skipped else ""
    flash("success", f"{int(mask.sum())} row(s) reset to pending.{note}")
    reset_editor()


def preview(fragment: str) -> None:
    """Render an e-mail body the way a mail client would: dark text on white, whatever the app theme."""
    st.html(f'<div style="background:#ffffff;color:#000000;padding:20px 24px;border-radius:6px;">{fragment}</div>')


def save_letter(letter_id: int, previous_status: str, fields: dict) -> None:
    if fields.get("status") == STATUS_PENDING and previous_status != STATUS_PENDING:
        fields.update(error_message="", sent_at="")
    try:
        store.update_row(letter_id, **fields)
    except QueueLockedError as exc:
        st.error(str(exc))
        return
    ss.letter_version += 1
    flash("success", f"Letter #{letter_id} saved.")
    st.rerun()


def run_import(keep_sent: bool) -> None:
    try:
        with st.spinner(f"Parsing {campaign.docx.name}…"):
            summary = docx_parser.import_docx(campaign, keep_sent=keep_sent)
    except Exception as exc:  # noqa: BLE001 - show parse problems in the UI
        flash("error", f"Import failed: {exc}")
        return
    flash(
        "success",
        f"Imported {summary['total']} letters (UA {summary['ua']}, EN {summary['en']}): "
        f"{summary['pending']} pending, {summary['missing_email']} without e-mail, "
        f"{summary['kept_sent']} kept as sent. Previous queue backed up to {campaign.backup_csv.name}.",
    )
    reset_editor()


# --------------------------------------------------------------------------- campaign

with st.sidebar:
    st.header("📬 Campaign")
    keys = list(CAMPAIGNS)
    ss.setdefault("campaign", DEFAULT_CAMPAIGN)
    campaign_key = st.radio(
        "Campaign", keys, key="campaign", label_visibility="collapsed",
        format_func=lambda k: CAMPAIGNS[k].label.split(" (")[0],
        help="Each campaign has its own document, queue and statuses. Only one sends at a time.",
    )
    st.caption(" · ".join(f"{CAMPAIGNS[k].label.split(' (')[0]}: {_campaign_summary(k)}" for k in keys))

campaign = CAMPAIGNS[campaign_key]
store = stores[campaign_key]
dispatcher = get_dispatcher(campaign)
if ss.get("editor_campaign") != campaign_key:  # a different queue: start the table over
    ss.editor_campaign = campaign_key
    reset_editor()

sending_now = active_campaign()  # whichever campaign holds the sender, in any browser session
busy = CAMPAIGNS[sending_now].label if sending_now and sending_now != campaign_key else ""

# --------------------------------------------------------------------------- first launch

if not store.exists():
    if campaign.docx.exists():
        run_import(keep_sent=True)
    else:
        st.warning(f"`{campaign.docx.name}` not found next to app.py, and there is no queue yet.")

smtp = config.load_settings()
queue = store.load()
label_for = row_label(queue)
ids = queue["id"].tolist()
first_pending = int(next(iter(queue.loc[queue["status"] == STATUS_PENDING, "id"]), ids[0] if ids else 0))

# --------------------------------------------------------------------------- sidebar

with st.sidebar:
    st.header("⏱️ Schedule")
    current = dispatcher.settings  # seed widgets from the worker so a page reload doesn't reset them
    ss.setdefault("interval_unit", "seconds")
    ss.setdefault("interval_s", int(current.interval_seconds))
    ss.setdefault("interval_m", round(current.interval_seconds / 60, 1))
    ss.setdefault("jitter", (int(current.jitter_min), int(current.jitter_max)))
    ss.setdefault("max_per_run", int(current.max_per_run))

    unit = st.radio("Interval unit", ["seconds", "minutes"], horizontal=True, key="interval_unit")
    if unit == "seconds":
        interval = float(st.number_input("Interval between e-mails (s)", 0, 86_400, step=10, key="interval_s"))
    else:
        interval = 60 * float(st.number_input("Interval between e-mails (min)", 0.0, 1_440.0, step=0.5, key="interval_m"))
    jitter = st.slider(
        "Random jitter ± (seconds)", 0, 120, key="jitter",
        help="Each delay is the interval plus or minus a random amount in this range, so sends don't arrive at a fixed rhythm.",
    )
    max_per_run = st.number_input(
        "Stop after N e-mails (0 = no limit)", 0, 5_000, step=10, key="max_per_run",
        help="Useful for staying under your provider's daily sending limit.",
    )
    send_settings = SendSettings(interval, jitter[0], jitter[1], int(max_per_run))
    dispatcher.update_settings(send_settings)
    low, high = send_settings.delay_range()
    st.caption(f"Delay between e-mails: **{fmt_duration(low)} – {fmt_duration(high)}** (minimum 5 s).")

    if st.checkbox("Enable scheduled start", key="sched_on", help="Press Start now, and sending begins later, e.g. while you're at school."):
        mode = st.radio("Scheduled start", SCHEDULE_MODES, horizontal=True, key="sched_mode", label_visibility="collapsed")
        if mode == SCHEDULE_MODES[0]:
            st.time_input("Start sending at", value=DEFAULT_START_TIME, step=300, key="sched_time")
        else:
            d1, d2 = st.columns([3, 2])
            d1.number_input("Start in", min_value=1, max_value=10_000, value=60, step=5, key="sched_delay")
            d2.selectbox("Unit", ["minutes", "hours"], key="sched_unit")
        target = scheduled_start()
        st.caption(
            f"After **Start**, sending begins at **{fmt_when(target)}** "
            f"(in {fmt_duration((target - datetime.now()).total_seconds())}). Keep this computer on, plugged in "
            "and awake (sleep off) with the app running; closing the browser tab is fine."
        )

    st.divider()
    st.header("🧪 Dry run")
    st.caption("Sends one letter (with its attachment) to a debug address. The queue status is not changed.")
    ss.setdefault("debug_email", smtp.debug_email or smtp.sender_email)
    st.text_input("Debug address", key="debug_email")
    if ids:
        if ss.get("test_row") not in ids:
            ss["test_row"] = first_pending
        st.selectbox("Letter to test", ids, key="test_row", format_func=label_for)
        if st.button("📨 Send test e-mail", width="stretch", disabled=not smtp.is_complete):
            with st.spinner("Sending test e-mail…"):
                ok, message = dispatcher.send_test(ss.test_row, ss.debug_email)
            (st.success if ok else st.error)(message)

    st.divider()
    st.header("📄 Import")
    keep_sent = st.checkbox(
        "Keep 'sent' status of letters already sent", value=True,
        help="Matched by recipient + subject, so re-importing never re-sends a letter that already went out.",
    )
    if st.button(f"🔄 Re-import {campaign.docx.name}", width="stretch",
                 disabled=dispatcher.is_active or not campaign.docx.exists()):
        if pending_edits():
            flash("warning", "Unsaved table edits were discarded by the import.")
        run_import(keep_sent)
        st.rerun()
    if dispatcher.is_active:
        st.caption("Stop the dispatcher to re-import.")

# --------------------------------------------------------------------------- header + SMTP banner

st.title(f"✉️ Email Dispatcher · {campaign.label.split(' (')[0]}")
files = attachment_files()
st.caption(
    f"Source: `{campaign.docx.name}` → `{campaign.queue_csv.relative_to(config.BASE_DIR).as_posix()}` · Attachments: "
    + (", ".join(f"`{f}`" for f in files) if files else "**none found in attachments/**")
)

if message := ss.pop("flash", None):
    getattr(st, message[0])(message[1])
if busy:
    st.info(f"**{busy}** is sending right now. Stop that campaign before starting this one.")
if not isinstance(dispatcher, Dispatcher):
    st.warning(
        "The app's code was updated while sending. This run continues with the previous version; "
        "the new version takes over when it finishes or you press Stop."
    )

banner, test_col = st.columns([5, 1], vertical_alignment="center")
with test_col:
    if st.button("🔌 Test SMTP", width="stretch", disabled=not smtp.is_complete):
        with st.spinner("Connecting…"):
            dispatcher.check_smtp()
with banner:
    if not config.ENV_FILE.exists():
        st.error("**No `.env` file.** Copy `.env.template` to `.env`, fill in your SMTP credentials, then reload.")
    elif not smtp.is_complete:
        st.error(f"**`.env` is incomplete:** missing {', '.join(smtp.missing_fields())}.")
    elif dispatcher.smtp_status:
        ok, text, checked_at = dispatcher.smtp_status
        (st.success if ok else st.error)(f"**SMTP {'OK' if ok else 'FAILED'}** · {text} · checked {checked_at}")
    else:
        st.info(f"SMTP **{smtp.describe()}** as **{smtp.sender_email}**, not tested yet.")

# --------------------------------------------------------------------------- live monitor

snap_at_run = dispatcher.snapshot()


@st.fragment(run_every=1 if snap_at_run["active"] else None)
def monitor() -> None:
    snap = dispatcher.snapshot()
    if (snap["active"], is_locked(snap)) != (snap_at_run["active"], is_locked(snap_at_run)):
        st.rerun()  # full rerun: lock/unlock the editor and switch live polling on/off

    counts = store.counts()
    total, sent, pending, errors = counts["total"], counts[STATUS_SENT], counts[STATUS_PENDING], counts[STATUS_ERROR]

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total", total)
    m2.metric("Pending", pending)
    m3.metric("Sent", sent)
    m4.metric("Errors", errors)
    eta = pending * max(dispatcher.settings.interval_seconds, 5) + snap["seconds_to_start"]
    m5.metric("ETA (pending)", fmt_duration(eta) if pending else "—")
    st.progress(sent / total if total else 0.0, text=f"Sent {sent} / {total}")

    state = snap["state"]
    if not snap["active"]:
        line = "⚪ **Idle**" + (f" · {snap['last_result']}" if snap["last_result"] else "")
    elif state == State.STOPPING:
        line = "⏹️ **Stopping…**"
    elif snap["paused"] and state == State.PAUSED and snap["scheduled_for"]:
        line = (
            f"⏸️ **Paused before the scheduled start** ({fmt_when(datetime.fromtimestamp(snap['scheduled_for']))}). "
            "Resume to keep waiting (sending starts right away if that time has passed), or Stop to cancel."
        )
    elif snap["paused"] and state == State.PAUSED:
        line = "⏸️ **Paused**. You can edit the queue now; press Resume to continue."
    elif snap["paused"]:
        line = "⏳ **Pausing** after the current e-mail…"
    elif state == State.SCHEDULED:
        line = (
            f"⏰ **Waiting until {fmt_when(datetime.fromtimestamp(snap['scheduled_for']))} to start sending** · "
            f"countdown **{fmt_countdown(snap['seconds_to_start'])}**"
        )
    elif state == State.SENDING:
        line = f"🟢 **Sending** {snap['current']}…"
    else:
        line = f"🟢 **Running** · next e-mail in **{fmt_duration(snap['seconds_to_next'])}**"
    if snap["active"] or snap["run_sent"] or snap["run_errors"]:
        line += f" · this run: {snap['run_sent']} sent, {snap['run_errors']} errors"
    if snap["active"] and snap["scheduled_for"]:
        st.info(line)  # a countdown is easy to miss as a plain line
    else:
        st.markdown(line)

    if ss.get("confirm_start"):
        limit = dispatcher.settings.max_per_run
        amount = f"up to **{limit}** of the **{pending}** pending" if limit and limit < pending else f"**{pending}**"
        target = scheduled_start()
        when = f" at **{fmt_when(target)}** (in {fmt_duration((target - datetime.now()).total_seconds())})" if target else ""
        st.warning(
            f"{'Schedule' if target else 'Start'} sending {amount} e-mail(s) to real recipients "
            f"from **{smtp.sender_email}**{when}? Tip: send a dry-run test first."
        )
        yes, no, _ = st.columns([1, 1, 3])
        if yes.button("✅ Yes, schedule" if target else "✅ Yes, start", type="primary", width="stretch"):
            ss.confirm_start = False
            flash("info", dispatcher.start(start_at=scheduled_start()))  # a delay counts from this click
            st.rerun()
        if no.button("Cancel", width="stretch"):
            ss.confirm_start = False
            st.rerun()
    else:
        b1, b2, b3, _ = st.columns([1, 1, 1, 2])
        if snap["active"] and snap["paused"]:
            if b1.button("▶️ Resume", type="primary", width="stretch"):
                dispatcher.start()
                st.rerun()
        elif snap["active"] and state == State.SCHEDULED:
            if b1.button("⏩ Start now", width="stretch", help="Skip the countdown and start sending immediately."):
                dispatcher.start_now()
                st.rerun()
        elif b1.button(
            "▶️ Start", type="primary", width="stretch",
            disabled=snap["active"] or not pending or not smtp.is_complete or bool(busy),
        ):
            ss.confirm_start = True
            st.rerun()
        if b2.button("⏸️ Pause", width="stretch", disabled=not snap["active"] or snap["paused"] or state == State.STOPPING):
            dispatcher.pause()
            st.rerun()
        if b3.button("⏹️ Stop", width="stretch", disabled=not snap["active"] or state == State.STOPPING):
            dispatcher.stop()
            st.rerun()

    log = dispatcher.recent_log(200)
    with st.expander(f"Activity log ({len(log)})", expanded=True):
        if log:
            st.dataframe(
                pd.DataFrame(
                    [(t, f"{LOG_ICONS.get(level, '')} {text}") for t, level, text in log],
                    columns=["Time", "Event"],
                ),
                hide_index=True,
                height=230,
                column_config={
                    "Time": st.column_config.TextColumn(width="small"),
                    "Event": st.column_config.TextColumn(width="large"),
                },
            )
        else:
            st.caption("Nothing yet. Events are also appended to data/activity.log.")


monitor()

# --------------------------------------------------------------------------- tabs

tab_queue, tab_letter, tab_check = st.tabs(["📋 Queue", "✏️ Letter editor & preview", "✅ Pre-send check"])
locked = is_locked(snap_at_run)

with tab_queue:
    if locked:
        st.info("The queue is read-only while sending. **Pause** to edit; changes are picked up when you resume.")

        @st.fragment(run_every=2)
        def live_table() -> None:
            st.dataframe(
                store.load(),
                hide_index=True,
                height=560,
                column_order=EDITOR_COLUMN_ORDER,
            )

        live_table()
    else:
        mtime = store.mtime()
        if "editor_base" not in ss or (ss.get("editor_mtime") != mtime and not pending_edits()):
            ss.editor_base = store.load()
            ss.editor_mtime = mtime
            ss.editor_version += 1
        elif ss.get("editor_mtime") != mtime:
            st.warning("data/queue.csv changed since you started editing. Saving merges only the cells you changed.")

        base = ss.editor_base
        edits = pending_edits()
        shown = st.pills(
            "Show", STATUSES, selection_mode="multi", default=STATUSES, key="status_filter",
            format_func=lambda s: f"{STATUS_ICONS[s]} {s} ({int((base['status'] == s).sum())})",
            disabled=bool(edits), help="Save or discard your edits to change the filter.",
        )
        view = base[base["status"].isin(shown or STATUSES)].reset_index(drop=True)
        ss.editor_view = view
        options = sorted(set(files) | set(base["attachment_filename"]) | {NO_ATTACHMENT})
        st.data_editor(
            view,
            key=editor_key(),
            hide_index=True,
            height=520,
            num_rows="dynamic",
            column_order=EDITOR_COLUMN_ORDER,
            disabled=["id", "body", "error_message", "sent_at"],
            column_config={
                "id": st.column_config.NumberColumn("ID", width="small"),
                "status": st.column_config.SelectboxColumn("Status", options=STATUSES, required=True, width="small"),
                "organization": st.column_config.TextColumn("Organization", width="medium"),
                "recipient_email": st.column_config.TextColumn(
                    "Recipient e-mail", width="medium", help="Several addresses: separate with commas."
                ),
                "subject": st.column_config.TextColumn("Subject", width="large"),
                "lang": st.column_config.SelectboxColumn("Lang", options=LANGS, required=True, width="small"),
                "attachment_filename": st.column_config.SelectboxColumn(
                    "Attachment", options=options, width="medium",
                    help=f"File in attachments/. Empty = default for the language, '{NO_ATTACHMENT}' = no attachment.",
                ),
                "body": st.column_config.TextColumn(
                    "Body (plain text)", width="large", help="Edit letter bodies (HTML and plain text) in the Letter editor tab."
                ),
                "error_message": st.column_config.TextColumn("Error", width="medium"),
                "sent_at": st.column_config.TextColumn("Sent at", width="small"),
            },
        )
        edits = pending_edits()
        c1, c2, c3, c4 = st.columns(4)
        c1.button("💾 Save changes", type="primary", width="stretch", disabled=not edits, on_click=save_editor_changes)
        c2.button("↩️ Discard changes", width="stretch", disabled=not edits, on_click=reset_editor)
        c3.button(
            "🔁 Reset errors → pending", width="stretch", on_click=reset_errors,
            disabled=bool(edits) or not (base["status"] == STATUS_ERROR).any(),
        )
        c4.download_button(
            "⬇️ Download queue.csv", data=queue.to_csv(index=False).encode("utf-8-sig"),
            file_name="queue.csv", mime="text/csv", width="stretch",
        )
        if edits:
            n = len(edits.get("edited_rows", {})) + len(edits.get("added_rows", [])) + len(edits.get("deleted_rows", []))
            st.caption(f"✏️ Unsaved changes in {n} row(s).")

with tab_letter:
    if not ids:
        st.info("The queue is empty.")
    else:
        if ss.get("letter_id") not in ids:
            ss["letter_id"] = first_pending
        letter_id = st.selectbox("Letter", ids, key="letter_id", format_func=label_for)
        row = queue[queue["id"] == letter_id].iloc[0].to_dict()
        text, fragment = message_bodies(row)  # as sent: {{date}} filled in
        stored_text, stored_html = stored_bodies(row)  # as saved: what the editors show
        version = f"{letter_id}_{ss.letter_version}"  # fresh widgets after every save
        details, summary = st.columns([3, 2], gap="large")

        with details:
            with st.form(f"letter_{version}"):
                recipient = st.text_input("Recipient e-mail", row["recipient_email"], help="Several addresses: separate with commas.")
                subject = st.text_input("Subject", row["subject"])
                f1, f2, f3 = st.columns([1, 2, 1])
                lang = f1.selectbox("Lang", LANGS, index=LANGS.index(row["lang"]) if row["lang"] in LANGS else 0)
                attachment = f2.text_input(
                    "Attachment(s)", row["attachment_filename"],
                    help=f"File name(s) in attachments/, separated by ';'. Empty = default for the language, '{NO_ATTACHMENT}' = no attachment.",
                )
                status = f3.selectbox("Status", STATUSES, index=STATUSES.index(row["status"]))
                submitted = st.form_submit_button("💾 Save details", type="primary", disabled=locked)
            if submitted:
                save_letter(letter_id, row["status"], {
                    "recipient_email": recipient.strip(),
                    "subject": " ".join(subject.split()),
                    "lang": lang,
                    "attachment_filename": attachment.strip(),
                    "status": status,
                })

        with summary:
            problems = validate_row(row)
            if row["status"] == STATUS_SENT:
                st.success(f"Sent {row['sent_at']}")
            elif problems:
                st.error("Not sendable yet:\n\n" + "\n".join(f"- {p}" for p in problems))
            else:
                st.success("Ready to send.")
            if row["error_message"]:
                st.caption(f"Last error: {row['error_message']}")

            attachment_info = []
            for name in attachment_names(row):
                path = resolve_attachment(name)
                attachment_info.append(f"📎 `{name}` ({path.stat().st_size / 1024:,.0f} KB)" if path else f"❌ `{name}` (missing)")
            st.markdown(
                f"**From:** {smtp.sender_name + ' ' if smtp.sender_name else ''}&lt;{smtp.sender_email or '?'}&gt;  \n"
                f"**To:** {', '.join(split_recipients(row['recipient_email'])) or '—'}  \n"
                f"**Subject:** {row['subject'] or '—'}  \n"
                f"**Attachments:** {', '.join(attachment_info) or 'none'}  \n"
                f"**Body:** HTML {len(fragment.encode('utf-8')) / 1024:,.1f} KB"
                + ("" if row["body_html"] else " (generated from the plain text)")
                + f" + plain-text fallback ({len(text):,} characters)"
            )
        if locked:
            st.caption("Pause the dispatcher to edit letters.")

        view = st.segmented_control("Body", BODY_VIEWS, default=BODY_VIEWS[0], key="body_view") or BODY_VIEWS[0]
        if view == BODY_VIEWS[0]:
            with st.container(border=True, height=640):
                preview(fragment)
            st.caption(
                "What recipients see in Gmail, Ukr.net, Outlook or Apple Mail. "
                "Mail clients that can't display HTML show the plain-text version instead."
                + (f" `{DATE_TOKEN}` shows today's date here and is filled with the sending date;" if DATE_TOKEN in stored_html else "")
                + (f" `{RECIPIENT_TOKEN}` shows the Recipient e-mail above." if RECIPIENT_TOKEN in stored_html else "")
            )

        elif view == BODY_VIEWS[1]:
            source_col, live_col = st.columns(2, gap="medium")
            with source_col:
                source = st.text_area(
                    "HTML source", stored_html, height=600, key=f"html_{version}",
                    help=f"{DATE_TOKEN} becomes the sending date, {RECIPIENT_TOKEN} the letter's recipient e-mail.",
                )
                sync = st.checkbox("Regenerate the plain-text version from this HTML", value=True, key=f"sync_{version}")
                if st.button("💾 Save HTML", type="primary", disabled=locked, key=f"save_html_{version}"):
                    if source.strip() == stored_html and not sync:
                        st.toast("No changes to save.")
                    else:
                        fields = {"body_html": source.strip()}
                        if sync:
                            fields["body"] = html_to_text(source)
                        save_letter(letter_id, row["status"], fields)
            with live_col:
                st.caption("Preview of the source on the left (updates when you click outside the editor or press Ctrl+Enter).")
                with st.container(border=True, height=600):
                    preview(message_bodies({**row, "body_html": source})[1])

        else:
            plain = st.text_area("Plain-text version", stored_text, height=560, key=f"text_{version}")
            rebuild = st.checkbox(
                "Also rebuild the HTML from this text (the Word formatting is lost)", value=False, key=f"rebuild_{version}"
            )
            st.caption(
                "Shown only by mail clients that don't display HTML; most recipients see the HTML version. "
                "To change what they see, edit the HTML source."
            )
            if st.button("💾 Save plain text", type="primary", disabled=locked, key=f"save_text_{version}"):
                plain = plain.replace("\r\n", "\n").strip()
                if plain == stored_text and not rebuild:
                    st.toast("No changes to save.")
                else:
                    fields = {"body": plain}
                    if rebuild:
                        fields["body_html"] = text_to_html(plain)
                    save_letter(letter_id, row["status"], fields)

with tab_check:
    pending_rows = queue[queue["status"] == STATUS_PENDING]
    issues = [
        {"id": r["id"], "organization": r["organization"], "recipient_email": r["recipient_email"], "problems": "; ".join(p)}
        for r in pending_rows.to_dict("records")
        if (p := validate_row(r))
    ]
    c1, c2, c3 = st.columns(3)
    c1.metric("Pending rows", len(pending_rows))
    c2.metric("Pending rows with problems", len(issues))
    c3.metric("Rows without e-mail", int((queue["recipient_email"] == "").sum()))
    if issues:
        st.error("These pending rows will be marked **error** when the dispatcher reaches them:")
        st.dataframe(pd.DataFrame(issues), hide_index=True)
    elif len(pending_rows):
        st.success("All pending rows pass validation (recipient, subject, body and attachment).")

    # Same address in several letters: often intentional (different departments), worth a look.
    address_rows: dict[str, list[str]] = {}
    for r in queue.to_dict("records"):
        for address in split_recipients(r["recipient_email"]):
            address_rows.setdefault(address.lower(), []).append(f"#{r['id']} ({r['status']})")
    dupes = {a: rows for a, rows in address_rows.items() if len(rows) > 1}
    if dupes:
        st.warning(f"{len(dupes)} address(es) appear in more than one letter:")
        st.dataframe(
            pd.DataFrame([{"address": a, "letters": ", ".join(rows)} for a, rows in dupes.items()]),
            hide_index=True,
        )

    st.subheader("Attachments")
    usage: dict[str, int] = {}
    for r in pending_rows.to_dict("records"):
        for name in attachment_names(r):
            usage[name] = usage.get(name, 0) + 1
    rows = []
    for name in sorted(set(files) | set(usage)):
        path = resolve_attachment(name)
        rows.append(
            {
                "file": name,
                "exists": "✅" if path else "❌ missing",
                "size": f"{path.stat().st_size / 1024:,.0f} KB" if path else "",
                "pending letters using it": usage.get(name, 0),
            }
        )
    st.dataframe(pd.DataFrame(rows), hide_index=True)
