"""
config.py — loads environment variables and exposes typed settings.
"""

import os
from zoneinfo import ZoneInfo

from dotenv import load_dotenv

load_dotenv()


def _require(key: str) -> str:
    value = os.getenv(key)
    if not value:
        raise EnvironmentError(
            f"Required environment variable '{key}' is missing. "
            "Copy .env.example to .env and fill in your values."
        )
    return value


# Bot token from @BotFather
BOT_TOKEN: str = _require("BOT_TOKEN")

# Admin user IDs (can run privileged commands)
_raw_admins = os.getenv("ADMIN_IDS", "")
ADMIN_IDS: set[int] = {
    int(uid.strip()) for uid in _raw_admins.split(",") if uid.strip().isdigit()
}

# Timezone for all timestamps
_tz_name = os.getenv("TIMEZONE", "UTC")
try:
    TIMEZONE = ZoneInfo(_tz_name)
except Exception:
    raise EnvironmentError(
        f"Unknown timezone '{_tz_name}'. Use a valid IANA name like 'Africa/Cairo'."
    )

# Data directory
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)

# Main CSV file — all attendance records live here
CSV_FILE = os.path.join(DATA_DIR, "attendance.csv")

# Excel export directory
EXPORT_DIR = os.path.join(DATA_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)
