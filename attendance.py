"""
attendance.py — attendance data layer.

Two stores are maintained in parallel:
  1. attendance.csv  — raw append-only log (all completed entries)
  2. master_attendance.xlsx — single print-ready master sheet; updated live
     on every /present (row added, Status=Present) and /leave (row filled).

CSV columns:
    User_ID, Username, Full_Name,
    Join_Date, Join_Time, Leave_Date, Leave_Time,
    Duration_Minutes, Session, Chat_ID

Master Excel columns (print sheet):
    No | User ID | Full Name | Date | Present Time | Leave Time | Duration | Status
"""

import os
from datetime import datetime, timezone
from typing import Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from config import CSV_FILE, MASTER_XLSX, TIMEZONE

# ── CSV columns ───────────────────────────────────────────────────────────────
COLUMNS = [
    "User_ID", "Username", "Full_Name",
    "Join_Date", "Join_Time",
    "Leave_Date", "Leave_Time",
    "Duration_Minutes", "Session", "Chat_ID",
]

# ── Master Excel columns (display order) ─────────────────────────────────────
MASTER_COLS = [
    "No", "User_ID", "Full_Name",
    "Date", "Present_Time", "Leave_Time",
    "Duration", "Status",
]

# In-memory active store
# { chat_id: { user_id: { username, name, join_time, session } } }
_active: dict[int, dict[int, dict]] = {}

# ── styling constants ─────────────────────────────────────────────────────────
_TITLE_FONT   = Font(name="Calibri", bold=True, size=14)
_HEADER_FONT  = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
_HEADER_FILL  = PatternFill("solid", fgColor="4472C4")   # blue
_PRESENT_FILL = PatternFill("solid", fgColor="E2EFDA")   # light green
_LEFT_FILL    = PatternFill("solid", fgColor="FCE4D6")   # light orange
_ALT_FILL     = PatternFill("solid", fgColor="F2F2F2")   # light grey
_CENTER       = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT_ALIGN   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

_THIN  = Side(style="thin",   color="AAAAAA")
_MED   = Side(style="medium", color="444444")
_THIN_BORDER   = Border(left=_THIN,  right=_THIN,  top=_THIN,  bottom=_THIN)
_HEADER_BORDER = Border(left=_MED,   right=_MED,   top=_MED,   bottom=_MED)


# ── helpers ───────────────────────────────────────────────────────────────────

def _now() -> datetime:
    return datetime.now(tz=timezone.utc).astimezone(TIMEZONE)


def _dur_str(minutes: float) -> str:
    total_secs = int(minutes * 60)
    h, rem = divmod(total_secs, 3600)
    m, s   = divmod(rem, 60)
    if h:
        return f"{h}h {m}m {s}s"
    return f"{m}m {s}s"


# ── CSV helpers ───────────────────────────────────────────────────────────────

def init_csv() -> None:
    if not os.path.exists(CSV_FILE):
        pd.DataFrame(columns=COLUMNS).to_csv(CSV_FILE, index=False)


def _read_csv() -> pd.DataFrame:
    init_csv()
    return pd.read_csv(CSV_FILE, dtype={"User_ID": "Int64", "Chat_ID": "Int64"})


def _append_csv(row: dict) -> None:
    init_csv()
    join_dt:  datetime = row["join_time"]
    leave_dt: datetime = row["leave_time"]
    entry = {
        "User_ID":          row["user_id"],
        "Username":         row["username"] or "",
        "Full_Name":        row["name"],
        "Join_Date":        join_dt.strftime("%Y-%m-%d"),
        "Join_Time":        join_dt.strftime("%H:%M:%S"),
        "Leave_Date":       leave_dt.strftime("%Y-%m-%d"),
        "Leave_Time":       leave_dt.strftime("%H:%M:%S"),
        "Duration_Minutes": round(row["duration_minutes"], 2),
        "Session":          row["session"],
        "Chat_ID":          row["chat_id"],
    }
    df = pd.read_csv(CSV_FILE)
    df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
    df.to_csv(CSV_FILE, index=False)


# ── Master Excel helpers ──────────────────────────────────────────────────────

def _apply_cell(cell, value, font=None, fill=None, alignment=None, border=None):
    cell.value = value
    if font:      cell.font      = font
    if fill:      cell.fill      = fill
    if alignment: cell.alignment = alignment
    if border:    cell.border    = border


def _rebuild_master() -> None:
    """
    Rebuild master_attendance.xlsx from scratch using the CSV as source of truth.
    Called once at bot start or after a session delete. Normally _upsert_master
    is used for incremental updates.
    """
    df = _read_csv()
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    _write_master_sheet(ws, df)
    _set_print_props(ws)
    wb.save(MASTER_XLSX)


def _write_master_sheet(ws, df: pd.DataFrame) -> None:
    """Write title row, header row, and all data rows to ws."""
    # ── Title ─────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value     = "Student / Member Attendance Summary Sheet"
    title_cell.font      = _TITLE_FONT
    title_cell.alignment = _CENTER
    ws.row_dimensions[1].height = 28

    # ── Header ────────────────────────────────────────────────────────────────
    headers = ["No", "User ID", "Full Name", "Date",
               "Present Time", "Leave Time", "Duration", "Status"]
    for col, hdr in enumerate(headers, 1):
        c = ws.cell(2, col, hdr)
        c.font      = _HEADER_FONT
        c.fill      = _HEADER_FILL
        c.alignment = _CENTER
        c.border    = _HEADER_BORDER
    ws.row_dimensions[2].height = 18

    # ── Data rows ─────────────────────────────────────────────────────────────
    for i, (_, row) in enumerate(df.iterrows(), 1):
        r = i + 2   # Excel row index (1=title, 2=header, 3+=data)

        status   = str(row.get("Status", "Left")) if "Status" in df.columns else "Left"
        is_pres  = status == "Present"
        row_fill = _PRESENT_FILL if is_pres else (_ALT_FILL if i % 2 == 0 else None)

        dur_raw  = row.get("Duration_Minutes", 0)
        dur_disp = "—" if is_pres else _dur_str(float(dur_raw) if pd.notna(dur_raw) else 0)

        vals = [
            i,
            int(row["User_ID"]) if pd.notna(row.get("User_ID")) else "",
            str(row.get("Full_Name", "")),
            str(row.get("Join_Date", "")),
            str(row.get("Join_Time", "")),
            str(row.get("Leave_Time", "")) if not is_pres else "—",
            dur_disp,
            status,
        ]
        for col, val in enumerate(vals, 1):
            c = ws.cell(r, col, val)
            c.alignment = _CENTER if col in (1, 4, 5, 6, 7, 8) else _LEFT_ALIGN
            c.border    = _THIN_BORDER
            if row_fill:
                c.fill = row_fill

        ws.row_dimensions[r].height = 16

    # ── Column widths ─────────────────────────────────────────────────────────
    widths = [6, 14, 28, 12, 14, 12, 14, 12]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


def _set_print_props(ws) -> None:
    """Configure the worksheet for clean printing."""
    ws.page_setup.orientation     = "landscape"
    ws.page_setup.fitToPage       = True
    ws.page_setup.fitToWidth      = 1
    ws.page_setup.fitToHeight     = 0          # unlimited pages tall
    ws.page_setup.paperSize       = ws.PAPERSIZE_A4
    ws.print_title_rows           = "1:2"      # repeat title+header on every page
    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.75,
        header=0.3, footer=0.3,
    )
    ws.sheet_view.showGridLines = True


def _upsert_master(user_id: int, name: str, join_dt: datetime,
                   leave_dt: Optional[datetime], duration_minutes: Optional[float],
                   session: str, status: str) -> None:
    """
    Open master_attendance.xlsx (or create it), then:
    - On /present  (status='Present') : append a new row with no leave info.
    - On /leave    (status='Left')    : find the matching Present row and fill it.
    Saves the file atomically.
    """
    # ── Load or create workbook ───────────────────────────────────────────────
    if os.path.exists(MASTER_XLSX):
        wb = load_workbook(MASTER_XLSX)
        ws = wb.active
        # Count existing data rows (skip title row 1 and header row 2)
        existing_rows = ws.max_row - 2
        if existing_rows < 0:
            existing_rows = 0
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Attendance"
        # Title row
        ws.merge_cells("A1:H1")
        t = ws["A1"]
        t.value, t.font, t.alignment = (
            "Student / Member Attendance Summary Sheet",
            _TITLE_FONT, _CENTER,
        )
        ws.row_dimensions[1].height = 28
        # Header row
        headers = ["No", "User ID", "Full Name", "Date",
                   "Present Time", "Leave Time", "Duration", "Status"]
        for col, hdr in enumerate(headers, 1):
            c = ws.cell(2, col, hdr)
            c.font, c.fill, c.alignment, c.border = (
                _HEADER_FONT, _HEADER_FILL, _CENTER, _HEADER_BORDER,
            )
        ws.row_dimensions[2].height = 18
        _set_print_props(ws)
        existing_rows = 0

    if status == "Present":
        # ── Append a new Present row ──────────────────────────────────────────
        row_num  = existing_rows + 3          # +1 title, +1 header, +1 new
        seq_no   = existing_rows + 1
        row_fill = _PRESENT_FILL

        vals = [
            seq_no,
            user_id,
            name,
            join_dt.strftime("%Y-%m-%d"),
            join_dt.strftime("%H:%M:%S"),
            "—",          # Leave Time not yet known
            "—",          # Duration not yet known
            "Present",
        ]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row_num, col, val)
            c.alignment = _CENTER if col in (1, 4, 5, 6, 7, 8) else _LEFT_ALIGN
            c.border    = _THIN_BORDER
            c.fill      = row_fill
        ws.row_dimensions[row_num].height = 16

    else:
        # ── Find existing Present row and fill in leave info ──────────────────
        target_row = None
        for r in range(3, ws.max_row + 1):
            cell_uid    = ws.cell(r, 2).value   # col B = User ID
            cell_date   = ws.cell(r, 4).value   # col D = Date
            cell_status = ws.cell(r, 8).value   # col H = Status
            if (str(cell_uid) == str(user_id)
                    and cell_date == join_dt.strftime("%Y-%m-%d")
                    and cell_status == "Present"):
                target_row = r
                break

        if target_row:
            dur_disp = _dur_str(duration_minutes) if duration_minutes is not None else "—"
            leave_str = leave_dt.strftime("%H:%M:%S") if leave_dt else "—"

            ws.cell(target_row, 6).value = leave_str    # Leave Time
            ws.cell(target_row, 7).value = dur_disp     # Duration
            ws.cell(target_row, 8).value = "Left"       # Status

            # Re-style the updated row
            row_fill = _LEFT_FILL
            for col in range(1, 9):
                c = ws.cell(target_row, col)
                c.fill   = row_fill
                c.border = _THIN_BORDER
        else:
            # Fallback: no matching Present row found — append a complete Left row
            row_num  = ws.max_row + 1
            seq_no   = row_num - 2
            dur_disp = _dur_str(duration_minutes) if duration_minutes is not None else "—"
            leave_str = leave_dt.strftime("%H:%M:%S") if leave_dt else "—"

            vals = [
                seq_no, user_id, name,
                join_dt.strftime("%Y-%m-%d"),
                join_dt.strftime("%H:%M:%S"),
                leave_str, dur_disp, "Left",
            ]
            for col, val in enumerate(vals, 1):
                c = ws.cell(row_num, col, val)
                c.alignment = _CENTER if col in (1, 4, 5, 6, 7, 8) else _LEFT_ALIGN
                c.border    = _THIN_BORDER
                c.fill      = _LEFT_FILL
            ws.row_dimensions[row_num].height = 16

    wb.save(MASTER_XLSX)


# ── public API ────────────────────────────────────────────────────────────────

def record_join(
    chat_id: int,
    session_label: str,
    user_id: int,
    name: str,
    username: Optional[str],
) -> datetime:
    """Store join in memory and add a 'Present' row to the master Excel."""
    join_time = _now()
    _active.setdefault(chat_id, {})[user_id] = {
        "username": username,
        "name": name,
        "join_time": join_time,
        "session": session_label,
    }
    # Write Present row immediately to master Excel
    _upsert_master(
        user_id=user_id, name=name,
        join_dt=join_time, leave_dt=None,
        duration_minutes=None,
        session=session_label, status="Present",
    )
    return join_time


def record_leave(
    chat_id: int,
    session_label: str,
    user_id: int,
) -> Optional[datetime]:
    """Pop from memory, append to CSV, and update the master Excel row to 'Left'."""
    entry = _active.get(chat_id, {}).pop(user_id, None)
    if entry is None:
        return None

    leave_time      = _now()
    join_time       = entry["join_time"]
    stored_session  = entry.get("session") or session_label
    duration_mins   = (leave_time - join_time).total_seconds() / 60

    # 1. Append to CSV log
    _append_csv({
        "user_id":          user_id,
        "username":         entry["username"],
        "name":             entry["name"],
        "join_time":        join_time,
        "leave_time":       leave_time,
        "duration_minutes": duration_mins,
        "session":          stored_session,
        "chat_id":          chat_id,
    })

    # 2. Update the master Excel row
    _upsert_master(
        user_id=user_id, name=entry["name"],
        join_dt=join_time, leave_dt=leave_time,
        duration_minutes=duration_mins,
        session=stored_session, status="Left",
    )
    return leave_time


def close_all_active(chat_id: int, session_label: str) -> int:
    user_ids = list(_active.get(chat_id, {}).keys())
    for uid in user_ids:
        record_leave(chat_id, session_label, uid)
    return len(user_ids)


def get_session_df(chat_id: int, session_label: str) -> pd.DataFrame:
    df = _read_csv()
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)
    mask = (df["Chat_ID"] == chat_id) & (df["Session"] == session_label)
    return df[mask].copy().reset_index(drop=True)


def list_sessions(chat_id: int) -> list[str]:
    df = _read_csv()
    if df.empty:
        return []
    return df[df["Chat_ID"] == chat_id]["Session"].dropna().unique().tolist()


def delete_session(chat_id: int, session_label: str) -> bool:
    df = _read_csv()
    mask = (df["Chat_ID"] == chat_id) & (df["Session"] == session_label)
    if not mask.any():
        return False
    df = df[~mask].reset_index(drop=True)
    df.to_csv(CSV_FILE, index=False)
    _rebuild_master()   # rebuild master from updated CSV
    return True


def get_active_participants(chat_id: int) -> list[dict]:
    return [
        {"user_id": uid, **info}
        for uid, info in _active.get(chat_id, {}).items()
    ]


def is_active(chat_id: int, user_id: int) -> bool:
    return user_id in _active.get(chat_id, {})
