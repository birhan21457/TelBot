"""
config.py — bot configuration.

Values are hardcoded below as defaults and can be overridden by a .env file.
If no .env is present the hardcoded defaults are used, so the bot starts
without needing any environment setup.
"""

import os
from zoneinfo import ZoneInfo

from dotenv import load_dotenv

load_dotenv()  # loads .env if it exists; no-op if it doesn't

# ── Hardcoded defaults (override via .env if needed) ─────────────────────────

_DEFAULT_TOKEN    = "8431291161:AAF0e2-zQwRK0Vg9WHJ19FNN5DCRsEVywiM"
_DEFAULT_ADMINS   = "8211053962"
_DEFAULT_TIMEZONE = "Africa/Nairobi"   # EAT (UTC+3)

# ── Bot token ─────────────────────────────────────────────────────────────────

BOT_TOKEN: str = os.getenv("BOT_TOKEN") or _DEFAULT_TOKEN

# ── Admin user IDs ────────────────────────────────────────────────────────────

_raw_admins = os.getenv("ADMIN_IDS") or _DEFAULT_ADMINS
ADMIN_IDS: set[int] = {
    int(uid.strip()) for uid in _raw_admins.split(",") if uid.strip().isdigit()
}

# ── Timezone ──────────────────────────────────────────────────────────────────

_tz_name = os.getenv("TIMEZONE") or _DEFAULT_TIMEZONE
try:
    TIMEZONE = ZoneInfo(_tz_name)
except Exception:
    print(f"[config] Unknown timezone '{_tz_name}', falling back to UTC.")
    TIMEZONE = ZoneInfo("UTC")

# ── Paths ─────────────────────────────────────────────────────────────────────

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)

CSV_FILE = os.path.join(DATA_DIR, "attendance.csv")

EXPORT_DIR = os.path.join(DATA_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)
