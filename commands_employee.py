"""
commands_employee.py – Employee-facing commands for SISWIT Attendance Bot.

Commands: /mystatus, /myprofile, /edit, /leave
"""

import logging
from datetime import datetime, timedelta

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import CommandHandler, ContextTypes

import config
from config import (
    load_staff,
    load_daily_log,
    load_leave_log,
    get_attendance_date,
    now,
    get_admin_ids,
)

logger = logging.getLogger(__name__)


def register_employee_commands(app) -> None:
    """Register all employee command handlers on the application."""
    app.add_handler(CommandHandler("mystatus", mystatus_command))
    app.add_handler(CommandHandler("myprofile", myprofile_command))
    app.add_handler(CommandHandler("edit", edit_command))
    app.add_handler(CommandHandler("leave", leave_command))


# ══════════════════════════════════════════════════════════════════════════════

async def mystatus_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/mystatus — Show the caller's attendance for the last 7 days."""
    user_id = update.effective_user.id
    staff = load_staff()
    daily_log = load_daily_log()
    leave_log = load_leave_log()

    # Reverse-lookup employee by linked telegram_id
    emp_id = None
    for eid, info in staff.items():
        if info.get("telegram_id") == user_id:
            emp_id = eid
            break

    if not emp_id:
        await update.message.reply_text(
            "❌ Your Telegram account is not linked to any employee ID.\n"
            "Submit attendance once with your EMP_ID to link automatically."
        )
        return

    emp_name = staff[emp_id]["name"]
    today = now().date()

    lines = [f"📊 *Weekly Status — {emp_name} ({emp_id})*\n"]
    for i in range(6, -1, -1):
        day = today - timedelta(days=i)
        ds = day.strftime("%Y-%m-%d")
        dn = day.strftime("%a")

        if ds in daily_log and emp_id in daily_log[ds]:
            entry = daily_log[ds][emp_id]
            status = "❌ Late" if entry.get("late") else "✅ On Time"
            lines.append(f"  {dn} {ds}: {status} — {entry.get('time', 'N/A')}")
        elif ds in leave_log and emp_id in leave_log.get(ds, {}):
            lines.append(f"  {dn} {ds}: 🏖 Leave")
        else:
            if day <= today:
                lines.append(f"  {dn} {ds}: ❌ Absent")
            else:
                lines.append(f"  {dn} {ds}: ⏳ Upcoming")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


# ══════════════════════════════════════════════════════════════════════════════

async def myprofile_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/myprofile — Display the caller's employee profile."""
    user_id = update.effective_user.id
    staff = load_staff()

    emp_id = None
    for eid, info in staff.items():
        if info.get("telegram_id") == user_id:
            emp_id = eid
            break

    if not emp_id:
        await update.message.reply_text(
            "❌ Your Telegram account is not linked to any employee ID."
        )
        return

    info = staff[emp_id]
    await update.message.reply_text(
        f"👤 *Employee Profile*\n\n"
        f"🆔 ID: `{emp_id}`\n"
        f"📛 Name: {info['name']}\n"
        f"🏢 Department: {info.get('dept', 'N/A')}\n"
        f"📱 Telegram ID: `{info.get('telegram_id', 'Not linked')}`",
        parse_mode="Markdown",
    )


# ══════════════════════════════════════════════════════════════════════════════

async def edit_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/edit EMP_ID New text — Request an edit to today's submission."""
    if not context.args or len(context.args) < 2:
        await update.message.reply_text("Usage: /edit EMP_ID New work description")
        return

    emp_id = context.args[0].upper()
    new_text = " ".join(context.args[1:])

    staff = load_staff()
    if emp_id not in staff:
        await update.message.reply_text(f"❌ Unknown employee: {emp_id}")
        return

    att_date = get_attendance_date()
    daily_log = load_daily_log()

    if att_date not in daily_log or emp_id not in daily_log.get(att_date, {}):
        await update.message.reply_text(
            f"❌ No submission found for {emp_id} on {att_date}. Nothing to edit."
        )
        return

    old_text = daily_log[att_date][emp_id].get("work", "N/A")
    edit_key = f"{emp_id}:{att_date}"
    config.pending_edits[edit_key] = {"new_text": new_text}

    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "✅ Approve", callback_data=f"edit_approve:{emp_id}:{att_date}",
                ),
                InlineKeyboardButton(
                    "❌ Reject", callback_data=f"edit_reject:{emp_id}:{att_date}",
                ),
            ]
        ]
    )

    msg = (
        f"✏️ *Edit Request*\n\n"
        f"Employee: {staff[emp_id]['name']} ({emp_id})\n"
        f"Date: {att_date}\n"
        f"Requested by: {update.effective_user.full_name}\n\n"
        f"*Old:* {old_text}\n"
        f"*New:* {new_text}\n\n"
        f"Approve or reject?"
    )

    for admin_id in get_admin_ids():
        try:
            await context.bot.send_message(
                admin_id, msg, reply_markup=keyboard, parse_mode="Markdown",
            )
        except Exception as exc:
            logger.warning("Failed to send edit request to admin %s: %s", admin_id, exc)

    await update.message.reply_text(
        f"📨 Edit request for {emp_id} sent to admin for approval."
    )


# ══════════════════════════════════════════════════════════════════════════════

async def leave_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/leave EMP_ID [today|tomorrow|DD-MM-YYYY] Reason"""
    if not context.args or len(context.args) < 3:
        await update.message.reply_text(
            "Usage: /leave EMP_ID [today|tomorrow|DD-MM-YYYY] Reason"
        )
        return

    emp_id = context.args[0].upper()
    date_arg = context.args[1].lower()
    reason = " ".join(context.args[2:])

    staff = load_staff()
    if emp_id not in staff:
        await update.message.reply_text(f"❌ Unknown employee: {emp_id}")
        return

    # Parse leave date
    today = now().date()
    if date_arg == "today":
        leave_date = today
    elif date_arg == "tomorrow":
        leave_date = today + timedelta(days=1)
    else:
        try:
            leave_date = datetime.strptime(date_arg, "%d-%m-%Y").date()
        except ValueError:
            await update.message.reply_text(
                "❌ Invalid date format. Use: today, tomorrow, or DD-MM-YYYY"
            )
            return

    leave_date_str = leave_date.strftime("%Y-%m-%d")
    leave_key = f"{emp_id}:{leave_date_str}"
    config.pending_leaves[leave_key] = {
        "reason": reason,
        "leave_date": leave_date_str,
    }

    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "✅ Approve",
                    callback_data=f"leave_approve:{emp_id}:{leave_date_str}",
                ),
                InlineKeyboardButton(
                    "❌ Reject",
                    callback_data=f"leave_reject:{emp_id}:{leave_date_str}",
                ),
            ]
        ]
    )

    msg = (
        f"🏖 *Leave Request*\n\n"
        f"Employee: {staff[emp_id]['name']} ({emp_id})\n"
        f"Department: {staff[emp_id].get('dept', 'N/A')}\n"
        f"Leave Date: {leave_date_str} ({leave_date.strftime('%A')})\n"
        f"Reason: {reason}\n"
        f"Requested by: {update.effective_user.full_name}\n\n"
        f"Approve or reject?"
    )

    for admin_id in get_admin_ids():
        try:
            await context.bot.send_message(
                admin_id, msg, reply_markup=keyboard, parse_mode="Markdown",
            )
        except Exception as exc:
            logger.warning("Failed to send leave request to admin %s: %s", admin_id, exc)

    await update.message.reply_text(
        f"📨 Leave request for {emp_id} on {leave_date_str} sent to admin for approval."
    )
