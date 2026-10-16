"""
attendance.py — CSV-based attendance store using pandas.

CSV schema (one row per completed session entry):
    User_ID          int       Telegram user ID
    Username         str       @handle or empty string
    Full_Name        str       First + Last name
    Join_Date        str       "YYYY-MM-DD"
    Join_Time        str       "HH:MM:SS"
    Leave_Date       str       "YYYY-MM-DD"
    Leave_Time       str       "HH:MM:SS"
    Duration_Minutes float     Rounded to 2 dp
    Session          str       Session label (e.g. "Math Class — 2024-10-16")
    Chat_ID          int       Telegram chat ID

Active (in-call) participants are held in memory in _active:
    _active[chat_id][user_id] = {
        "username": str | None,
        "name": str,
        "join_time": datetime,   # full datetime with tz
        "session": str,
    }
Once a user leaves, the completed row is appended to the CSV.
"""

import os
from datetime import datetime, timezone
from typing import Optional

import pandas as pd

from config import CSV_FILE, TIMEZONE

# ── CSV columns in exact order ────────────────────────────────────────────────
COLUMNS = [
    "User_ID",
    "Username",
    "Full_Name",
    "Join_Date",
    "Join_Time",
    "Leave_Date",
    "Leave_Time",
    "Duration_Minutes",
    "Session",
    "Chat_ID",
]

# In-memory store for participants currently in the call.
# Structure: { chat_id: { user_id: { username, name, join_time, session } } }
_active: dict[int, dict[int, dict]] = {}


# ── CSV helpers ───────────────────────────────────────────────────────────────

def init_csv() -> None:
    """Create the CSV file with headers if it does not exist."""
    if not os.path.exists(CSV_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_csv(CSV_FILE, index=False)


def _read() -> pd.DataFrame:
    init_csv()
    return pd.read_csv(CSV_FILE, dtype={"User_ID": "Int64", "Chat_ID": "Int64"})


def _append_row(row: dict) -> None:
    """
    Append a single completed attendance row to the CSV.
    join_dt and leave_dt are tz-aware datetime objects.
    """
    init_csv()
    join_dt:  datetime = row["join_time"]
    leave_dt: datetime = row["leave_time"]

    new_entry = {
        "User_ID":           row["user_id"],
        "Username":          row["username"] or "",
        "Full_Name":         row["name"],
        "Join_Date":         join_dt.strftime("%Y-%m-%d"),
        "Join_Time":         join_dt.strftime("%H:%M:%S"),
        "Leave_Date":        leave_dt.strftime("%Y-%m-%d"),
        "Leave_Time":        leave_dt.strftime("%H:%M:%S"),
        "Duration_Minutes":  round(row["duration_minutes"], 2),
        "Session":           row["session"],
        "Chat_ID":           row["chat_id"],
    }

    df = pd.read_csv(CSV_FILE)
    df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
    df.to_csv(CSV_FILE, index=False)


def _now() -> datetime:
    return datetime.now(tz=timezone.utc).astimezone(TIMEZONE)


# ── public API ────────────────────────────────────────────────────────────────

def record_join(
    chat_id: int,
    session_label: str,
    user_id: int,
    name: str,
    username: Optional[str],
) -> datetime:
    """
    Record that user_id joined the session.
    Stores in memory; the CSV row is only written on leave.
    Returns the join timestamp.
    """
    join_time = _now()
    _active.setdefault(chat_id, {})[user_id] = {
        "username": username,
        "name": name,
        "join_time": join_time,
        "session": session_label,
    }
    return join_time


def record_leave(
    chat_id: int,
    session_label: str,
    user_id: int,
) -> Optional[datetime]:
    """
    Record that user_id left the session.
    Always uses the session label stored at join time so the CSV row
    is guaranteed to match the session the user joined under.
    Returns the leave timestamp, or None if no active entry existed.
    """
    entry = _active.get(chat_id, {}).pop(user_id, None)
    if entry is None:
        return None

    leave_time = _now()
    join_time: datetime = entry["join_time"]
    stored_session: str = entry.get("session") or session_label
    duration_minutes = (leave_time - join_time).total_seconds() / 60

    _append_row({
        "user_id":          user_id,
        "username":         entry["username"],
        "name":             entry["name"],
        "join_time":        join_time,
        "leave_time":       leave_time,
        "duration_minutes": duration_minutes,
        "session":          stored_session,
        "chat_id":          chat_id,
    })
    return leave_time


def close_all_active(chat_id: int, session_label: str) -> int:
    """
    Force-close every participant still marked as active (e.g. on video_chat_ended).
    Returns the number of entries closed.
    """
    user_ids = list(_active.get(chat_id, {}).keys())
    for uid in user_ids:
        record_leave(chat_id, session_label, uid)
    return len(user_ids)


def get_session_df(chat_id: int, session_label: str) -> pd.DataFrame:
    """Return a DataFrame of all completed rows for a session."""
    df = _read()
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)
    mask = (df["Chat_ID"] == chat_id) & (df["Session"] == session_label)
    return df[mask].copy().reset_index(drop=True)


def list_sessions(chat_id: int) -> list[str]:
    """Return all distinct session labels for a chat."""
    df = _read()
    if df.empty:
        return []
    return df[df["Chat_ID"] == chat_id]["Session"].dropna().unique().tolist()


def delete_session(chat_id: int, session_label: str) -> bool:
    """Delete all rows for a session. Returns True if any rows were removed."""
    df = _read()
    mask = (df["Chat_ID"] == chat_id) & (df["Session"] == session_label)
    if not mask.any():
        return False
    df = df[~mask].reset_index(drop=True)
    df.to_csv(CSV_FILE, index=False)
    return True


def get_active_participants(chat_id: int) -> list[dict]:
    """Return participants currently in the call (in-memory only)."""
    return [
        {"user_id": uid, **info}
        for uid, info in _active.get(chat_id, {}).items()
    ]


def is_active(chat_id: int, user_id: int) -> bool:
    """Return True if user_id is currently marked as present in chat_id."""
    return user_id in _active.get(chat_id, {})
