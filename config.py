"""
config.py – Configuration, constants, and persistence helpers for SISWIT Attendance Bot.
"""

import os
import json
import logging
from pathlib import Path
from datetime import datetime, timedelta

import pytz
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)

# ── Telegram ─────────────────────────────────────────────────────────────────
BOT_TOKEN: str = os.getenv("BOT_TOKEN", "")
OWNER_CHAT_ID: int = int(os.getenv("OWNER_CHAT_ID", "0"))
HR_CHAT_ID: int = int(os.getenv("HR_CHAT_ID", "0"))

# ── Files ────────────────────────────────────────────────────────────────────
EXCEL_FILE: str = os.getenv("EXCEL_FILE", "attendance.xlsx")
STAFF_FILE: str = "staff.json"
DAILY_LOG_FILE: str = "daily_log.json"
LEAVE_LOG_FILE: str = "leave_log.json"
DEADLINE_FILE: str = "deadline.json"

# ── Google Sheets ────────────────────────────────────────────────────────────
GOOGLE_SHEET_ID: str = os.getenv("GOOGLE_SHEET_ID", "")
GOOGLE_CREDS_FILE: str = os.getenv("GOOGLE_CREDS_FILE", "credentials.json")
GOOGLE_CREDS_JSON: str = os.getenv("GOOGLE_CREDS_JSON", "")

# ── Timezone ─────────────────────────────────────────────────────────────────
TIMEZONE_STR: str = os.getenv("TIMEZONE", "Asia/Kolkata")
TIMEZONE = pytz.timezone(TIMEZONE_STR)

# ── Defaults ─────────────────────────────────────────────────────────────────
DEFAULT_DEADLINE: str = os.getenv("SUBMISSION_DEADLINE", "11:00")
ATTENDANCE_CUTOFF_HOUR: int = 13  # 1:00 PM

# ── Excel / Sheets Headers ──────────────────────────────────────────────────
ATTENDANCE_HEADERS: list[str] = [
    "S.No.", "Date", "Day", "Employee Name", "Emp ID", "Department",
    "Submit Time", "Work Report", "Source",
    "Telegram User", "Revision",
]

LEAVE_HEADERS: list[str] = [
    "Sr No", "Employee ID", "Name", "Department", "Leave Date",
    "Reason", "Approved By", "Status", "Leave #", "Deduction",
]

# ── Runtime State (in-memory, not persisted across restarts) ─────────────────
group_chat_id: int | None = None

pending_allows: dict = {}            # "EMP_ID:DATE" → True
allow_counts: dict = {}              # "EMP_ID:DATE" → int
pending_edits: dict = {}             # "EMP_ID:DATE" → {"new_text": str}
pending_leaves: dict = {}            # "EMP_ID:DATE" → {"reason": str, "leave_date": str}
approved_resubmissions: dict = {}    # "EMP_ID:DATE" → True


# ══════════════════════════════════════════════════════════════════════════════
# Persistence Helpers
# ══════════════════════════════════════════════════════════════════════════════

def _load_json(filepath: str, default=None):
    """Load a JSON file; return *default* on any error."""
    if default is None:
        default = {}
    try:
        if Path(filepath).exists():
            with open(filepath, "r", encoding="utf-8") as fh:
                return json.load(fh)
    except Exception as exc:
        logger.error("Failed to load %s: %s", filepath, exc)
    return default


def _save_json(filepath: str, data) -> None:
    """Atomically write *data* to a JSON file (write-tmp then rename)."""
    try:
        tmp = filepath + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp, filepath)
    except Exception as exc:
        logger.error("Failed to save %s: %s", filepath, exc)


# ── Staff ────────────────────────────────────────────────────────────────────

def load_staff() -> dict:
    """Return the staff registry ``{EMP_ID: {name, dept, telegram_id}}``."""
    return _load_json(STAFF_FILE, {})


def save_staff(staff: dict) -> None:
    _save_json(STAFF_FILE, staff)


# ── Daily Log ────────────────────────────────────────────────────────────────

def load_daily_log() -> dict:
    """Return ``{YYYY-MM-DD: {EMP_ID: {time, work, late, …}}}``."""
    return _load_json(DAILY_LOG_FILE, {})


def save_daily_log(log: dict) -> None:
    _save_json(DAILY_LOG_FILE, log)


# ── Leave Log ────────────────────────────────────────────────────────────────

def load_leave_log() -> dict:
    """Return ``{YYYY-MM-DD: {EMP_ID: {approved_by, reason}}, _monthly_counts: …}``."""
    return _load_json(LEAVE_LOG_FILE, {"_monthly_counts": {}})


def save_leave_log(log: dict) -> None:
    _save_json(LEAVE_LOG_FILE, log)


# ── Deadline ─────────────────────────────────────────────────────────────────

def load_deadline() -> str:
    """Return the current submission deadline as ``"HH:MM"``."""
    data = _load_json(DEADLINE_FILE, {"deadline": DEFAULT_DEADLINE})
    return data.get("deadline", DEFAULT_DEADLINE)


def save_deadline(deadline: str) -> None:
    _save_json(DEADLINE_FILE, {"deadline": deadline})


# ── HR Chat ID (persisted so it survives restarts) ───────────────────────────

_HR_FILE = "hr_chat_id.json"


def load_hr_chat_id() -> int:
    data = _load_json(_HR_FILE, {"hr_chat_id": HR_CHAT_ID})
    return int(data.get("hr_chat_id", HR_CHAT_ID))


def save_hr_chat_id(chat_id: int) -> None:
    global HR_CHAT_ID
    HR_CHAT_ID = chat_id
    _save_json(_HR_FILE, {"hr_chat_id": chat_id})


# ══════════════════════════════════════════════════════════════════════════════
# Convenience Helpers
# ══════════════════════════════════════════════════════════════════════════════

def now() -> datetime:
    """Current datetime in the configured timezone."""
    return datetime.now(TIMEZONE)


def get_attendance_date() -> str:
    """
    Compute the attendance date with the **1:00 PM cutoff**.

    * Before 1 PM  → previous calendar day
    * 1 PM or after → current day

    Returns ``"YYYY-MM-DD"``.
    """
    current = now()
    if current.hour < ATTENDANCE_CUTOFF_HOUR:
        target = current - timedelta(days=1)
    else:
        target = current
    return target.strftime("%Y-%m-%d")


def is_admin(chat_id: int) -> bool:
    """``True`` if *chat_id* belongs to the Owner or HR."""
    hr = load_hr_chat_id()
    return chat_id in (OWNER_CHAT_ID, hr)


def get_admin_ids() -> list[int]:
    """Return a de-duplicated list of admin chat-IDs (Owner + HR)."""
    hr = load_hr_chat_id()
    ids = [OWNER_CHAT_ID]
    if hr and hr != OWNER_CHAT_ID:
        ids.append(hr)
    return ids
