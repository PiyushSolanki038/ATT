"""
commands_admin.py – Admin / Owner commands for SISWIT Attendance Bot.

Commands registered here are available to OWNER_CHAT_ID and HR_CHAT_ID only
(except /staff which lists employees for anyone in the group).
"""

import logging
from datetime import datetime, timedelta
from pathlib import Path

from telegram import Update
from telegram.ext import CommandHandler, ContextTypes

import config
from config import (
    OWNER_CHAT_ID,
    EXCEL_FILE,
    TIMEZONE,
    load_staff,
    save_staff,
    load_daily_log,
    load_leave_log,
    get_attendance_date,
    now,
    is_admin,
    get_admin_ids,
    load_deadline,
    save_deadline,
    save_hr_chat_id,
)
import excel_handler

logger = logging.getLogger(__name__)


def register_admin_commands(app) -> None:
    """Register all admin command handlers on the application."""
    app.add_handler(CommandHandler("staff", staff_command))
    app.add_handler(CommandHandler("addstaff", addstaff_command))
    app.add_handler(CommandHandler("removestaff", removestaff_command))
    app.add_handler(CommandHandler("report", report_command))
    app.add_handler(CommandHandler("absent", absent_command))
    app.add_handler(CommandHandler("late", late_command))
    app.add_handler(CommandHandler("history", history_command))
    app.add_handler(CommandHandler("weeklyreport", weeklyreport_command))
    app.add_handler(CommandHandler("monthly", monthly_command))
    app.add_handler(CommandHandler("export", export_command))
    app.add_handler(CommandHandler("broadcast", broadcast_command))
    app.add_handler(CommandHandler("deadline", deadline_command))
    app.add_handler(CommandHandler("sethr", sethr_command))
    app.add_handler(CommandHandler("announce", announce_command))
    app.add_handler(CommandHandler("dm", dm_command))
    app.add_handler(CommandHandler("remind", remind_command))
    app.add_handler(CommandHandler("warning", warning_command))


# ══════════════════════════════════════════════════════════════════════════════
# Staff Management
# ══════════════════════════════════════════════════════════════════════════════

async def staff_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/staff — List all registered employees."""
    staff = load_staff()
    if not staff:
        await update.message.reply_text("📋 No employees registered. Use /addstaff to add.")
        return

    lines = ["📋 *Registered Employees*\n"]
    for emp_id, info in sorted(staff.items()):
        linked = "✅" if info.get("telegram_id") else "❌"
        lines.append(
            f"  `{emp_id}` — {info['name']} | {info.get('dept', 'N/A')} | Linked: {linked}"
        )
    lines.append(f"\nTotal: {len(staff)} employees")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def addstaff_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/addstaff EMP_ID Name | Department"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return
    if not context.args or len(context.args) < 2:
        await update.message.reply_text("Usage: /addstaff EMP_ID Name | Department")
        return

    emp_id = context.args[0].upper()
    rest = " ".join(context.args[1:])

    if "|" in rest:
        parts = rest.split("|", 1)
        name = parts[0].strip()
        dept = parts[1].strip()
    else:
        name = rest.strip()
        dept = "General"

    staff = load_staff()
    staff[emp_id] = {"name": name, "dept": dept, "telegram_id": None}
    save_staff(staff)

    await update.message.reply_text(
        f"✅ Added employee:\n"
        f"  ID: `{emp_id}`\n"
        f"  Name: {name}\n"
        f"  Department: {dept}",
        parse_mode="Markdown",
    )


async def removestaff_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/removestaff EMP_ID"""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return
    if not context.args:
        await update.message.reply_text("Usage: /removestaff EMP_ID")
        return

    emp_id = context.args[0].upper()
    staff = load_staff()

    if emp_id not in staff:
        await update.message.reply_text(f"❌ Employee {emp_id} not found.")
        return

    name = staff[emp_id]["name"]
    del staff[emp_id]
    save_staff(staff)
    await update.message.reply_text(f"✅ Removed employee: {name} ({emp_id})")


# ══════════════════════════════════════════════════════════════════════════════
# Reports
# ══════════════════════════════════════════════════════════════════════════════

async def report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/report — Today's submitted / absent / leave breakdown."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return

    att_date = get_attendance_date()
    staff = load_staff()
    daily_log = load_daily_log()
    leave_log = load_leave_log()

    day_entries = daily_log.get(att_date, {})
    day_leaves = leave_log.get(att_date, {})

    submitted, absent, on_leave = [], [], []

    for emp_id, info in sorted(staff.items()):
        if emp_id in day_entries:
            entry = day_entries[emp_id]
            submitted.append(f"  ✅ {info['name']} ({emp_id}) — {entry.get('time', 'N/A')}")
        elif emp_id in day_leaves:
            on_leave.append(f"  🏖 {info['name']} ({emp_id})")
        else:
            absent.append(f"  ❌ {info['name']} ({emp_id})")

    total = len(staff)
    pct = round(len(submitted) / total * 100) if total else 0

    lines = [
        f"📊 *Daily Report — {att_date}*\n",
        f"Submission Rate: {len(submitted)}/{total} ({pct}%)\n",
    ]
    if submitted:
        lines.append(f"*Submitted ({len(submitted)}):*")
        lines.extend(submitted)
    if on_leave:
        lines.append(f"\n*On Leave ({len(on_leave)}):*")
        lines.extend(on_leave)
    if absent:
        lines.append(f"\n*Absent ({len(absent)}):*")
        lines.extend(absent)

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def absent_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/absent — Quick absent list for today."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return

    att_date = get_attendance_date()
    staff = load_staff()
    daily_log = load_daily_log()
    leave_log = load_leave_log()

    day_entries = daily_log.get(att_date, {})
    day_leaves = leave_log.get(att_date, {})

    absent_list = []
    for emp_id, info in sorted(staff.items()):
        if emp_id not in day_entries and emp_id not in day_leaves:
            absent_list.append(f"  ❌ {info['name']} ({emp_id})")

    if not absent_list:
        await update.message.reply_text(f"✅ All employees have submitted for {att_date}!")
    else:
        msg = f"📋 *Absent List — {att_date}*\n\n" + "\n".join(absent_list)
        msg += f"\n\nTotal absent: {len(absent_list)}/{len(staff)}"
        await update.message.reply_text(msg, parse_mode="Markdown")


async def late_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/late — Late vs on-time breakdown."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return

    att_date = get_attendance_date()
    staff = load_staff()
    daily_log = load_daily_log()
    day_entries = daily_log.get(att_date, {})

    late_list, ontime_list = [], []
    for emp_id, entry in sorted(day_entries.items()):
        emp_name = staff.get(emp_id, {}).get("name", emp_id)
        t = entry.get("time", "N/A")
        if entry.get("late"):
            late_list.append(f"  ❌ {emp_name} ({emp_id}) — {t}")
        else:
            ontime_list.append(f"  ✅ {emp_name} ({emp_id}) — {t}")

    deadline = load_deadline()
    lines = [f"⏰ *Punctuality Report — {att_date}*", f"Deadline: {deadline}\n"]
    if ontime_list:
        lines.append(f"*On Time ({len(ontime_list)}):*")
        lines.extend(ontime_list)
    if late_list:
        lines.append(f"\n*Late ({len(late_list)}):*")
        lines.extend(late_list)
    if not ontime_list and not late_list:
        lines.append("No submissions yet.")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def history_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/history EMP_ID — Last 7 days detail (sent to admin DM)."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return
    if not context.args:
        await update.message.reply_text("Usage: /history EMP_ID")
        return

    emp_id = context.args[0].upper()
    staff = load_staff()
    if emp_id not in staff:
        await update.message.reply_text(f"❌ Unknown employee: {emp_id}")
        return

    emp_name = staff[emp_id]["name"]
    daily_log = load_daily_log()
    leave_log = load_leave_log()
    today = now().date()

    lines = [f"📜 *History — {emp_name} ({emp_id})*\n"]
    for i in range(6, -1, -1):
        day = today - timedelta(days=i)
        ds = day.strftime("%Y-%m-%d")
        dn = day.strftime("%a")

        if ds in daily_log and emp_id in daily_log[ds]:
            entry = daily_log[ds][emp_id]
            status = "❌ LATE" if entry.get("late") else "✅ On Time"
            resub = " 🔄" if entry.get("is_resubmission") else ""
            lines.append(
                f"  {dn} {ds}: {status} [{entry.get('time', '')}]{resub}\n"
                f"    📋 {entry.get('work', 'N/A')}"
            )
        elif ds in leave_log and emp_id in leave_log.get(ds, {}):
            reason = leave_log[ds][emp_id].get("reason", "N/A")
            lines.append(f"  {dn} {ds}: 🏖 Leave — {reason}")
        else:
            lines.append(f"  {dn} {ds}: ❌ Absent")

    try:
        await context.bot.send_message(
            update.effective_user.id, "\n".join(lines), parse_mode="Markdown",
        )
        if update.effective_chat.type in ("group", "supergroup"):
            await update.message.reply_text("📨 History sent to your DM.")
    except Exception:
        await update.message.reply_text(
            "❌ Failed to send DM. Start a private chat with the bot first."
        )


async def weeklyreport_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/weeklyreport — Grid of ✅/❌/🏖 for all employees (DM)."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return

    staff = load_staff()
    daily_log = load_daily_log()
    leave_log = load_leave_log()
    today = now().date()
    dates = [today - timedelta(days=i) for i in range(6, -1, -1)]

    day_hdrs = " | ".join(d.strftime("%a")[:2] for d in dates)
    sep = "-+-".join(["--"] * 7)

    lines = [
        "📊 *Weekly Attendance Report*",
        f"Week: {dates[0].strftime('%d/%m')} – {dates[-1].strftime('%d/%m/%Y')}\n",
        "```",
        f"{'Name':<15} | {day_hdrs}",
        f"{'-' * 15}-+-{sep}",
    ]

    for emp_id, info in sorted(staff.items()):
        name = info["name"][:15]
        statuses = []
        for d in dates:
            ds = d.strftime("%Y-%m-%d")
            if ds in daily_log and emp_id in daily_log[ds]:
                statuses.append("✅")
            elif ds in leave_log and emp_id in leave_log.get(ds, {}):
                statuses.append("🏖")
            else:
                statuses.append("❌")
        lines.append(f"{name:<15} | {' | '.join(statuses)}")

    lines.append("```")

    try:
        await context.bot.send_message(
            update.effective_user.id, "\n".join(lines), parse_mode="Markdown",
        )
        if update.effective_chat.type in ("group", "supergroup"):
            await update.message.reply_text("📨 Weekly report sent to your DM.")
    except Exception:
        await update.message.reply_text(
            "❌ Failed to send DM. Start a private chat with the bot first."
        )


async def monthly_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/monthly — Per-employee monthly stats with progress bar (DM)."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return

    staff = load_staff()
    daily_log = load_daily_log()
    leave_log = load_leave_log()
    current = now()
    first_day = current.replace(day=1).date()
    today = current.date()
    working_days = (today - first_day).days + 1

    lines = [f"📊 *Monthly Report — {current.strftime('%B %Y')}*\n"]

    for emp_id, info in sorted(staff.items()):
        present = late = leaves = 0
        for offset in range(working_days):
            d = first_day + timedelta(days=offset)
            ds = d.strftime("%Y-%m-%d")
            if ds in daily_log and emp_id in daily_log[ds]:
                present += 1
                if daily_log[ds][emp_id].get("late"):
                    late += 1
            elif ds in leave_log and emp_id in leave_log.get(ds, {}):
                leaves += 1

        absent = working_days - present - leaves
        pct = round(present / working_days * 100) if working_days else 0
        bar_len = 10
        filled = round(pct / 100 * bar_len)
        bar = "█" * filled + "░" * (bar_len - filled)

        lines.append(
            f"*{info['name']}* ({emp_id})\n"
            f"  [{bar}] {pct}%\n"
            f"  ✅ Present: {present} | ❌ Late: {late} | "
            f"🏖 Leave: {leaves} | ⛔ Absent: {absent}"
        )

    try:
        await context.bot.send_message(
            update.effective_user.id, "\n".join(lines), parse_mode="Markdown",
        )
        if update.effective_chat.type in ("group", "supergroup"):
            await update.message.reply_text("📨 Monthly report sent to your DM.")
    except Exception:
        await update.message.reply_text(
            "❌ Failed to send DM. Start a private chat with the bot first."
        )


# ══════════════════════════════════════════════════════════════════════════════
# Export, Broadcast, Deadline, SetHR
# ══════════════════════════════════════════════════════════════════════════════

async def export_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/export — Send Excel file + Google Sheet link to admin DM."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return

    user_id = update.effective_user.id

    if Path(EXCEL_FILE).exists():
        try:
            with open(EXCEL_FILE, "rb") as f:
                await context.bot.send_document(
                    user_id, document=f, filename=EXCEL_FILE,
                    caption="📊 Attendance Excel File",
                )
        except Exception as exc:
            logger.error("Failed to send Excel: %s", exc)
            await context.bot.send_message(user_id, "❌ Failed to send Excel file.")
    else:
        await context.bot.send_message(user_id, "ℹ️ No Excel file generated yet.")

    sheet_url = excel_handler.get_google_sheet_url()
    if sheet_url:
        await context.bot.send_message(user_id, f"📊 Google Sheet:\n{sheet_url}")

    if update.effective_chat.type in ("group", "supergroup"):
        await update.message.reply_text("📨 Export sent to your DM.")


async def broadcast_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/broadcast message — Send a message to the group."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return
    if not context.args:
        await update.message.reply_text("Usage: /broadcast Your message")
        return

    message = " ".join(context.args)

    if config.group_chat_id:
        try:
            await context.bot.send_message(
                config.group_chat_id,
                f"📢 *Broadcast*\n\n{message}",
                parse_mode="Markdown",
            )
            await update.message.reply_text("✅ Broadcast sent.")
        except Exception as exc:
            await update.message.reply_text(f"❌ Failed to broadcast: {exc}")
    else:
        await update.message.reply_text(
            "❌ Group chat not detected yet. Send a message in the group first."
        )


async def deadline_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/deadline [HH:MM] — View or set submission deadline."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return

    if not context.args:
        current = load_deadline()
        await update.message.reply_text(
            f"⏰ Current deadline: *{current}*", parse_mode="Markdown",
        )
        return

    raw = context.args[0]
    try:
        h, m = map(int, raw.split(":"))
        if not (0 <= h <= 23 and 0 <= m <= 59):
            raise ValueError
        formatted = f"{h:02d}:{m:02d}"
    except (ValueError, AttributeError):
        await update.message.reply_text("❌ Invalid format. Use HH:MM (e.g., 11:00)")
        return

    save_deadline(formatted)
    await update.message.reply_text(
        f"✅ Deadline updated to *{formatted}*", parse_mode="Markdown",
    )


async def sethr_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/sethr CHAT_ID — Owner-only: change the HR recipient."""
    if update.effective_user.id != OWNER_CHAT_ID:
        await update.message.reply_text("❌ Owner only command.")
        return
    if not context.args:
        await update.message.reply_text("Usage: /sethr CHAT_ID")
        return

    try:
        hr_id = int(context.args[0])
    except ValueError:
        await update.message.reply_text("❌ Invalid chat ID. Must be a number.")
        return

    save_hr_chat_id(hr_id)
    await update.message.reply_text(
        f"✅ HR chat ID updated to `{hr_id}`", parse_mode="Markdown",
    )


# ══════════════════════════════════════════════════════════════════════════════
# Private-chat admin utilities
# ══════════════════════════════════════════════════════════════════════════════

async def announce_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/announce message — Post an announcement to the group (DM only)."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return
    if update.effective_chat.type != "private":
        await update.message.reply_text("ℹ️ Use this command in a private chat with the bot.")
        return
    if not context.args:
        await update.message.reply_text("Usage: /announce Your announcement")
        return

    message = " ".join(context.args)
    if config.group_chat_id:
        try:
            await context.bot.send_message(
                config.group_chat_id,
                f"📣 *Announcement*\n\n{message}",
                parse_mode="Markdown",
            )
            await update.message.reply_text("✅ Announcement sent to group.")
        except Exception as exc:
            await update.message.reply_text(f"❌ Failed: {exc}")
    else:
        await update.message.reply_text("❌ Group chat not detected yet.")


async def dm_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/dm EMP_ID message — Send a private message to an employee."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return
    if not context.args or len(context.args) < 2:
        await update.message.reply_text("Usage: /dm EMP_ID Your message")
        return

    emp_id = context.args[0].upper()
    message = " ".join(context.args[1:])
    staff = load_staff()

    if emp_id not in staff:
        await update.message.reply_text(f"❌ Unknown employee: {emp_id}")
        return

    tid = staff[emp_id].get("telegram_id")
    if not tid:
        await update.message.reply_text(
            f"❌ {emp_id} doesn't have a linked Telegram account."
        )
        return

    try:
        await context.bot.send_message(
            tid, f"💬 *Message from Admin*\n\n{message}", parse_mode="Markdown",
        )
        await update.message.reply_text(
            f"✅ Message sent to {staff[emp_id]['name']} ({emp_id})."
        )
    except Exception as exc:
        await update.message.reply_text(f"❌ Failed to send message: {exc}")


async def remind_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/remind — Send a reminder to every employee who hasn't submitted."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return

    att_date = get_attendance_date()
    staff = load_staff()
    daily_log = load_daily_log()
    leave_log = load_leave_log()

    day_entries = daily_log.get(att_date, {})
    day_leaves = leave_log.get(att_date, {})

    pending, sent, failed = [], 0, 0

    for emp_id, info in staff.items():
        if emp_id not in day_entries and emp_id not in day_leaves:
            tid = info.get("telegram_id")
            if tid:
                try:
                    await context.bot.send_message(
                        tid,
                        f"⏰ *Reminder*\n\n"
                        f"Hi {info['name']}, you haven't submitted your attendance "
                        f"for {att_date} yet.\n"
                        f"Please submit in the group: `{emp_id} Your work description`",
                        parse_mode="Markdown",
                    )
                    sent += 1
                except Exception:
                    failed += 1
            pending.append(f"{info['name']} ({emp_id})")

    bullets = "\n".join(f"  • {p}" for p in pending) if pending else "  (none)"
    await update.message.reply_text(
        f"📨 Reminders sent.\n"
        f"Pending: {len(pending)} | Sent: {sent} | Failed: {failed}\n\n"
        f"Pending employees:\n{bullets}",
    )


async def warning_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/warning EMP_ID reason — Send an official warning to an employee."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("❌ Admin only command.")
        return
    if not context.args or len(context.args) < 2:
        await update.message.reply_text("Usage: /warning EMP_ID Reason")
        return

    emp_id = context.args[0].upper()
    reason = " ".join(context.args[1:])
    staff = load_staff()

    if emp_id not in staff:
        await update.message.reply_text(f"❌ Unknown employee: {emp_id}")
        return

    tid = staff[emp_id].get("telegram_id")
    if not tid:
        await update.message.reply_text(
            f"❌ {emp_id} doesn't have a linked Telegram account."
        )
        return

    try:
        await context.bot.send_message(
            tid,
            f"⚠️ *Official Warning*\n\n"
            f"Dear {staff[emp_id]['name']},\n\n"
            f"This is an official warning regarding:\n{reason}\n\n"
            f"Please take necessary action.",
            parse_mode="Markdown",
        )
        await update.message.reply_text(
            f"✅ Warning sent to {staff[emp_id]['name']} ({emp_id})."
        )
    except Exception as exc:
        await update.message.reply_text(f"❌ Failed to send warning: {exc}")
