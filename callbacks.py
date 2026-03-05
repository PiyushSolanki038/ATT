"""
callbacks.py – Inline-keyboard callback handlers for SISWIT Attendance Bot.

Handles Approve / Reject actions for:
  • Allow (re-submission)
  • Edit
  • Leave
"""

import asyncio
import logging
from datetime import datetime

from telegram import Update
from telegram.ext import CallbackQueryHandler, ContextTypes

import config
from config import (
    load_staff,
    load_daily_log,
    save_daily_log,
    load_leave_log,
    save_leave_log,
    is_admin,
)
import excel_handler

logger = logging.getLogger(__name__)


def register_callbacks(app) -> None:
    """Register all callback-query handlers on the application."""
    app.add_handler(CallbackQueryHandler(allow_callback, pattern=r"^allow_"))
    app.add_handler(CallbackQueryHandler(edit_callback, pattern=r"^edit_"))
    app.add_handler(CallbackQueryHandler(leave_callback, pattern=r"^leave_"))


# ══════════════════════════════════════════════════════════════════════════════
# Allow (re-submission) callbacks
# ══════════════════════════════════════════════════════════════════════════════

async def allow_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if not is_admin(query.from_user.id):
        await query.edit_message_text("❌ You are not authorized to approve requests.")
        return

    parts = query.data.split(":")  # e.g. "allow_approve:EMP001:2026-03-05"
    if len(parts) != 3:
        await query.edit_message_text("❌ Invalid callback data.")
        return

    action, emp_id, att_date = parts
    staff = load_staff()
    emp_name = staff.get(emp_id, {}).get("name", emp_id)
    approver = query.from_user.full_name

    if action == "allow_approve":
        # Remove existing entry from in-memory log so a new one can be saved
        daily_log = load_daily_log()
        if att_date in daily_log and emp_id in daily_log[att_date]:
            del daily_log[att_date][emp_id]
            save_daily_log(daily_log)

        # Flag the next submission as an approved re-submission
        config.approved_resubmissions[f"{emp_id}:{att_date}"] = True

        await query.edit_message_text(
            f"✅ *Re-submission Approved*\n\n"
            f"Employee: {emp_name} ({emp_id})\n"
            f"Date: {att_date}\n"
            f"Approved by: {approver}\n\n"
            f"Previous entry removed. Employee can now re-submit.",
            parse_mode="Markdown",
        )

        # Notify group
        if config.group_chat_id:
            try:
                await context.bot.send_message(
                    config.group_chat_id,
                    f"✅ Re-submission approved for {emp_name} ({emp_id}) on {att_date}.\n"
                    f"Please submit your updated work report now.",
                )
            except Exception as exc:
                logger.warning("Failed to notify group about allow: %s", exc)

    elif action == "allow_reject":
        config.pending_allows.pop(f"{emp_id}:{att_date}", None)
        await query.edit_message_text(
            f"❌ *Re-submission Rejected*\n\n"
            f"Employee: {emp_name} ({emp_id})\n"
            f"Date: {att_date}\n"
            f"Rejected by: {approver}",
            parse_mode="Markdown",
        )


# ══════════════════════════════════════════════════════════════════════════════
# Edit callbacks
# ══════════════════════════════════════════════════════════════════════════════

async def edit_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if not is_admin(query.from_user.id):
        await query.edit_message_text("❌ You are not authorized.")
        return

    parts = query.data.split(":")
    if len(parts) != 3:
        await query.edit_message_text("❌ Invalid callback data.")
        return

    action, emp_id, att_date = parts
    staff = load_staff()
    emp_name = staff.get(emp_id, {}).get("name", emp_id)
    approver = query.from_user.full_name
    edit_key = f"{emp_id}:{att_date}"

    if action == "edit_approve":
        edit_data = config.pending_edits.pop(edit_key, None)
        if not edit_data:
            await query.edit_message_text("❌ Edit request expired or already handled.")
            return

        new_text = edit_data["new_text"]

        # Update in-memory log
        daily_log = load_daily_log()
        if att_date in daily_log and emp_id in daily_log[att_date]:
            daily_log[att_date][emp_id]["work"] = new_text
            save_daily_log(daily_log)

        # Update Excel
        excel_handler.update_entry_in_excel(emp_id, att_date, new_text)

        # Update Google Sheets (async)
        async def _gs_edit():
            try:
                await excel_handler.update_entry_in_google_sheets(emp_id, att_date, new_text)
            except Exception as exc:
                logger.error("GSheets edit update failed (bg): %s", exc)
        asyncio.create_task(_gs_edit())

        await query.edit_message_text(
            f"✅ *Edit Approved*\n\n"
            f"Employee: {emp_name} ({emp_id})\n"
            f"Date: {att_date}\n"
            f"New text: {new_text}\n"
            f"Approved by: {approver}",
            parse_mode="Markdown",
        )

        if config.group_chat_id:
            try:
                await context.bot.send_message(
                    config.group_chat_id,
                    f"✏️ Edit approved for {emp_name} ({emp_id}) on {att_date}.",
                )
            except Exception as exc:
                logger.warning("Failed to notify group about edit: %s", exc)

    elif action == "edit_reject":
        config.pending_edits.pop(edit_key, None)
        await query.edit_message_text(
            f"❌ *Edit Rejected*\n\n"
            f"Employee: {emp_name} ({emp_id})\n"
            f"Date: {att_date}\n"
            f"Rejected by: {approver}",
            parse_mode="Markdown",
        )


# ══════════════════════════════════════════════════════════════════════════════
# Leave callbacks
# ══════════════════════════════════════════════════════════════════════════════

async def leave_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if not is_admin(query.from_user.id):
        await query.edit_message_text("❌ You are not authorized.")
        return

    parts = query.data.split(":")
    if len(parts) != 3:
        await query.edit_message_text("❌ Invalid callback data.")
        return

    action, emp_id, leave_date_str = parts
    staff = load_staff()
    emp_info = staff.get(emp_id, {})
    emp_name = emp_info.get("name", emp_id)
    dept = emp_info.get("dept", "N/A")
    approver = query.from_user.full_name
    leave_key = f"{emp_id}:{leave_date_str}"

    if action == "leave_approve":
        leave_data = config.pending_leaves.pop(leave_key, None)
        if not leave_data:
            await query.edit_message_text("❌ Leave request expired or already handled.")
            return

        reason = leave_data["reason"]

        # ── Update leave_log ─────────────────────────────────────────────
        leave_log = load_leave_log()
        leave_log.setdefault(leave_date_str, {})
        leave_log[leave_date_str][emp_id] = {
            "approved_by": approver,
            "reason": reason,
        }

        # Increment monthly counter
        leave_dt = datetime.strptime(leave_date_str, "%Y-%m-%d")
        month_key = leave_dt.strftime("%Y-%m")
        leave_log.setdefault("_monthly_counts", {})
        leave_log["_monthly_counts"].setdefault(month_key, {})
        current_count = leave_log["_monthly_counts"][month_key].get(emp_id, 0) + 1
        leave_log["_monthly_counts"][month_key][emp_id] = current_count
        save_leave_log(leave_log)

        # ── Save to Excel ────────────────────────────────────────────────
        excel_handler.save_leave_to_excel(
            emp_id, emp_name, dept, leave_date_str, reason, approver, current_count,
        )

        # ── Save to Google Sheets (async) ────────────────────────────────
        async def _gs_leave():
            try:
                await excel_handler.save_leave_to_google_sheets(
                    emp_id, emp_name, dept, leave_date_str, reason, approver, current_count,
                )
            except Exception as exc:
                logger.error("GSheets leave save failed (bg): %s", exc)
        asyncio.create_task(_gs_leave())

        # Deduction info (only shown in admin DM, NOT in group)
        deduction_info = ""
        if current_count > 3:
            extra = current_count - 3
            deduction_info = f"\n💰 Deduction: ₹{extra * 500} ({extra} extra leave(s))"

        await query.edit_message_text(
            f"✅ *Leave Approved*\n\n"
            f"Employee: {emp_name} ({emp_id})\n"
            f"Leave Date: {leave_date_str}\n"
            f"Reason: {reason}\n"
            f"Approved by: {approver}\n"
            f"Leave #{current_count} this month{deduction_info}",
            parse_mode="Markdown",
        )

        # Group notification (no financial details)
        if config.group_chat_id:
            try:
                await context.bot.send_message(
                    config.group_chat_id,
                    f"🏖 Leave approved for {emp_name} ({emp_id}) on {leave_date_str}.",
                )
            except Exception as exc:
                logger.warning("Failed to notify group about leave: %s", exc)

    elif action == "leave_reject":
        config.pending_leaves.pop(leave_key, None)
        await query.edit_message_text(
            f"❌ *Leave Rejected*\n\n"
            f"Employee: {emp_name} ({emp_id})\n"
            f"Leave Date: {leave_date_str}\n"
            f"Rejected by: {approver}",
            parse_mode="Markdown",
        )
