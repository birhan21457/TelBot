"""
bot.py — entry point.

Run:
    python bot.py
"""

import logging

from telegram.ext import (
    Application,
    ChatMemberHandler,
    CommandHandler,
    MessageHandler,
    filters,
)

from config import BOT_TOKEN
from handlers import (
    cmd_deletesession,
    cmd_export,
    cmd_leave,
    cmd_present,
    cmd_report,
    cmd_sessions,
    cmd_setsession,
    cmd_start,
    cmd_whoisin,
    on_chat_member_updated,
    on_video_chat_ended,
    on_video_chat_participants_invited,
    on_video_chat_started,
)

logging.basicConfig(
    format="%(asctime)s | %(levelname)-8s | %(name)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    level=logging.INFO,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger(__name__)


def main() -> None:
    app = Application.builder().token(BOT_TOKEN).build()

    # ── Video-chat service messages ───────────────────────────────────────────
    app.add_handler(MessageHandler(
        filters.StatusUpdate.VIDEO_CHAT_STARTED,
        on_video_chat_started,
    ))
    app.add_handler(MessageHandler(
        filters.StatusUpdate.VIDEO_CHAT_ENDED,
        on_video_chat_ended,
    ))
    app.add_handler(MessageHandler(
        filters.StatusUpdate.VIDEO_CHAT_PARTICIPANTS_INVITED,
        on_video_chat_participants_invited,
    ))

    # ── ChatMember updates — primary join/leave source ────────────────────────
    # Requires the bot to be a group admin so Telegram sends these events.
    app.add_handler(ChatMemberHandler(
        on_chat_member_updated,
        ChatMemberHandler.CHAT_MEMBER,
    ))

    # ── Commands ──────────────────────────────────────────────────────────────
    app.add_handler(CommandHandler("start",          cmd_start))
    app.add_handler(CommandHandler("present",        cmd_present))
    app.add_handler(CommandHandler("leave",          cmd_leave))
    app.add_handler(CommandHandler("setsession",     cmd_setsession))
    app.add_handler(CommandHandler("report",         cmd_report))
    app.add_handler(CommandHandler("export",         cmd_export))
    app.add_handler(CommandHandler("sessions",       cmd_sessions))
    app.add_handler(CommandHandler("whoisin",        cmd_whoisin))
    app.add_handler(CommandHandler("deletesession",  cmd_deletesession))

    logger.info("Bot is running — polling for updates…")
    app.run_polling(
        allowed_updates=[
            "message",
            "chat_member",      # required for join/leave detection
            "my_chat_member",
        ],
        drop_pending_updates=True,
    )


if __name__ == "__main__":
    main()
