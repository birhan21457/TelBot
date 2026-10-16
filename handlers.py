"""
handlers.py — Telegram update handlers and bot commands.

Video-chat events detected:
  • video_chat_started                  → announces tracking is active
  • video_chat_ended                    → force-closes all open entries
  • video_chat_participants_invited     → records invited users as joined
  • ChatMemberUpdated (chat_member)     → primary join/leave detection

Session naming:
  Auto-label = "<Chat Title> — YYYY-MM-DD"
  Admins can override with /setsession <label>.
"""

import html
import logging
import os
from datetime import datetime, timezone
from typing import Optional

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import ContextTypes

import attendance
from config import ADMIN_IDS, TIMEZONE
from reports import build_text_report, build_excel_report

logger = logging.getLogger(__name__)

# Per-chat active session label (in-memory).
_active_sessions: dict[int, str] = {}


# ── helpers ───────────────────────────────────────────────────────────────────

def _now_local() -> datetime:
    return datetime.now(tz=timezone.utc).astimezone(TIMEZONE)


def _auto_label(chat_id: int, context: ContextTypes.DEFAULT_TYPE) -> str:
    date_str = _now_local().strftime("%Y-%m-%d")
    title = context.bot_data.get(f"title_{chat_id}", f"Chat {chat_id}")
    return f"{title} — {date_str}"


def _session(chat_id: int, context: ContextTypes.DEFAULT_TYPE) -> str:
    return _active_sessions.get(chat_id) or _auto_label(chat_id, context)


def _full_name(user) -> str:
    parts = [user.first_name or "", user.last_name or ""]
    return " ".join(p for p in parts if p).strip() or f"User {user.id}"


def _is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


def _cache_title(chat, context: ContextTypes.DEFAULT_TYPE) -> None:
    if chat and chat.title:
        context.bot_data[f"title_{chat.id}"] = chat.title


# ── video-chat lifecycle ──────────────────────────────────────────────────────

async def on_video_chat_started(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    chat = update.effective_chat
    if not chat:
        return
    _cache_title(chat, context)
    label = _session(chat.id, context)
    attendance.init_csv()
    logger.info("Video chat started | chat=%s | session='%s'", chat.id, label)
    await update.effective_message.reply_text(
        f"🎙 <b>Video chat started</b>\n"
        f"📋 Session: <code>{html.escape(label)}</code>\n\n"
        f"Attendance tracking is <b>active</b>.\n"
        f"Admins can rename the session with /setsession &lt;label&gt;.",
        parse_mode=ParseMode.HTML,
    )


async def on_video_chat_ended(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    chat = update.effective_chat
    if not chat:
        return
    label = _session(chat.id, context)
    closed = attendance.close_all_active(chat.id, label)

    ended = update.effective_message.video_chat_ended
    duration_secs: int = ended.duration if ended else 0
    mins, secs = divmod(duration_secs, 60)

    logger.info("Video chat ended | chat=%s | closed %d open entries", chat.id, closed)
    await update.effective_message.reply_text(
        f"🔴 <b>Video chat ended</b>\n"
        f"📋 Session: <code>{html.escape(label)}</code>\n"
        f"⏱ Duration: {mins}m {secs}s\n\n"
        f"Use /report to see the attendance summary.",
        parse_mode=ParseMode.HTML,
    )


async def on_video_chat_participants_invited(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    chat = update.effective_chat
    msg = update.effective_message
    if not chat or not msg:
        return
    invited = msg.video_chat_participants_invited
    if not invited or not invited.users:
        return

    label = _session(chat.id, context)
    for user in invited.users:
        join_time = attendance.record_join(
            chat.id, label, user.id,
            _full_name(user),
            f"@{user.username}" if user.username else None,
        )
        logger.info("INVITED JOIN | %s (%s) | %s", _full_name(user), user.id, join_time)


# ── chat member updates (primary join/leave source) ───────────────────────────

async def on_chat_member_updated(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """
    Fires on every chat member status change.
    Detects video-chat participation via the `video_chat_participant` flag
    on ChatMember objects (requires bot to be group admin).
    """
    result = update.chat_member
    if not result:
        return

    chat = result.chat
    user = result.new_chat_member.user
    old  = result.old_chat_member
    new  = result.new_chat_member

    _cache_title(chat, context)
    label = _session(chat.id, context)
    name  = _full_name(user)
    uname = f"@{user.username}" if user.username else None

    old_in_vc = getattr(old, "video_chat_participant", False) or False
    new_in_vc = getattr(new, "video_chat_participant", False) or False

    if not old_in_vc and new_in_vc:
        # ── joined ────────────────────────────────────────────────────────────
        join_time = attendance.record_join(chat.id, label, user.id, name, uname)
        logger.info("JOIN  | %-25s (id=%s) | %s | session='%s'",
                    name, user.id, join_time.strftime("%H:%M:%S"), label)
        await context.bot.send_message(
            chat_id=chat.id,
            text=(
                f"✅ <b>{html.escape(name)}</b>"
                f"{' ' + html.escape(uname) if uname else ''} "
                f"joined the video chat\n"
                f"🕐 {join_time.strftime('%H:%M:%S')}"
            ),
            parse_mode=ParseMode.HTML,
        )

    elif old_in_vc and not new_in_vc:
        # ── left ──────────────────────────────────────────────────────────────
        leave_time = attendance.record_leave(chat.id, label, user.id)
        left_str = leave_time.strftime("%H:%M:%S") if leave_time else "unknown"
        logger.info("LEAVE | %-25s (id=%s) | %s | session='%s'",
                    name, user.id, left_str, label)
        await context.bot.send_message(
            chat_id=chat.id,
            text=(
                f"👋 <b>{html.escape(name)}</b>"
                f"{' ' + html.escape(uname) if uname else ''} "
                f"left the video chat\n"
                f"🕐 {left_str}"
            ),
            parse_mode=ParseMode.HTML,
        )


# ── commands ──────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "👋 <b>Attendance Bot</b>\n\n"
        "I automatically track who joins and leaves video chats in this group "
        "and record their times in a CSV file.\n\n"
        "<b>Commands:</b>\n"
        "/report — attendance summary for the current session\n"
        "/export — download attendance as an Excel file\n"
        "/sessions — list all recorded sessions\n"
        "/whoisin — show who is currently in the call\n"
        "/setsession &lt;label&gt; — set a custom session name (admins)\n"
        "/deletesession — delete current session data (admins)\n",
        parse_mode=ParseMode.HTML,
    )


async def cmd_setsession(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.effective_user
    chat = update.effective_chat
    if not _is_admin(user.id):
        await update.message.reply_text("⛔ Only admins can set the session name.")
        return
    if not context.args:
        await update.message.reply_text("Usage: /setsession <label>")
        return
    label = " ".join(context.args).strip()
    _active_sessions[chat.id] = label
    await update.message.reply_text(
        f"✅ Session label set to: <code>{html.escape(label)}</code>",
        parse_mode=ParseMode.HTML,
    )


async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    chat = update.effective_chat
    label = _session(chat.id, context)
    df = attendance.get_session_df(chat.id, label)
    if df.empty:
        await update.message.reply_text(
            f"No completed entries yet for session:\n"
            f"<code>{html.escape(label)}</code>\n\n"
            f"(Entries are saved when participants leave the call.)",
            parse_mode=ParseMode.HTML,
        )
        return
    report = build_text_report(label, df)
    # Telegram message limit is 4096 chars; split if needed
    for chunk in _split(report, 4096):
        await update.message.reply_text(chunk, parse_mode=ParseMode.HTML)


async def cmd_export(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    chat = update.effective_chat
    label = _session(chat.id, context)
    df = attendance.get_session_df(chat.id, label)
    if df.empty:
        await update.message.reply_text("No completed entries to export yet.")
        return
    filepath = build_excel_report(chat.id, label, df)
    with open(filepath, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename=os.path.basename(filepath),
            caption=f"📊 Attendance — {html.escape(label)}",
        )


async def cmd_sessions(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    chat = update.effective_chat
    sessions = attendance.list_sessions(chat.id)
    if not sessions:
        await update.message.reply_text("No sessions recorded yet.")
        return
    lines = "\n".join(f"• <code>{html.escape(s)}</code>" for s in sessions)
    await update.message.reply_text(
        f"<b>Recorded sessions ({len(sessions)}):</b>\n{lines}",
        parse_mode=ParseMode.HTML,
    )


async def cmd_whoisin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    chat = update.effective_chat
    active = attendance.get_active_participants(chat.id)
    if not active:
        await update.message.reply_text("Nobody is currently marked as in the call.")
        return
    now = _now_local()
    lines = []
    for p in active:
        upart = f" ({html.escape(p['username'])})" if p.get("username") else ""
        elapsed_secs = int((now - p["join_time"]).total_seconds())
        m, s = divmod(elapsed_secs, 60)
        lines.append(
            f"• <b>{html.escape(p['name'])}</b>{upart} — "
            f"joined {p['join_time'].strftime('%H:%M:%S')} ({m}m {s}s ago)"
        )
    await update.message.reply_text(
        f"<b>Currently in the call ({len(active)}):</b>\n" + "\n".join(lines),
        parse_mode=ParseMode.HTML,
    )


async def cmd_deletesession(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.effective_user
    chat = update.effective_chat
    if not _is_admin(user.id):
        await update.message.reply_text("⛔ Only admins can delete session data.")
        return
    label = _session(chat.id, context)
    deleted = attendance.delete_session(chat.id, label)
    if deleted:
        await update.message.reply_text(
            f"🗑 Session <code>{html.escape(label)}</code> deleted from CSV.",
            parse_mode=ParseMode.HTML,
        )
    else:
        await update.message.reply_text(
            f"No data found for session: <code>{html.escape(label)}</code>",
            parse_mode=ParseMode.HTML,
        )


# ── utility ───────────────────────────────────────────────────────────────────

def _split(text: str, limit: int) -> list[str]:
    """Split a long string into chunks ≤ limit characters on newline boundaries."""
    if len(text) <= limit:
        return [text]
    chunks, current = [], []
    length = 0
    for line in text.splitlines(keepends=True):
        if length + len(line) > limit and current:
            chunks.append("".join(current))
            current, length = [], 0
        current.append(line)
        length += len(line)
    if current:
        chunks.append("".join(current))
    return chunks
