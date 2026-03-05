"""
bot.py – Main entry point and message handler for SISWIT Employee Attendance Bot.

Registers all command / message / callback handlers and starts polling.
"""

import asyncio
import logging
import sys
from datetime import datetime

from telegram import Update, BotCommand, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

import config
from config import (
    BOT_TOKEN,
    OWNER_CHAT_ID,
    load_staff,
    save_staff,
    load_daily_log,
    save_daily_log,
    get_attendance_date,
    now,
    is_admin,
    get_admin_ids,
    load_deadline,
)
import excel_handler
from commands_admin import register_admin_commands
from commands_employee import register_employee_commands
from callbacks import register_callbacks

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    level=logging.INFO,
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# /start  &  /help
# ══════════════════════════════════════════════════════════════════════════════

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Welcome to SISWIT Attendance Bot!*\n\n"
        "Submit your daily attendance in the group:\n"
        "`EMP_ID Your work description`\n\n"
        "Type /help for all commands.",
        parse_mode="Markdown",
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    base = (
        "📋 *SISWIT Attendance Bot — Commands*\n\n"
        "*Everyone:*\n"
        "• Submit: `EMP_ID Your work description`\n"
        "• /mystatus — Your weekly attendance status\n"
        "• /myprofile — Your profile info\n"
        "• /edit `EMP_ID New text` — Request edit\n"
        "• /leave `EMP_ID [today|tomorrow|DD-MM-YYYY] Reason`\n"
        "• /allow `EMP_ID` — Request re-submission\n"
    )

    admin = ""
    if is_admin(user_id):
        admin = (
            "\n*Admin (Group):*\n"
            "• /staff — List employees\n"
            "• /addstaff `EMP_ID Name | Dept`\n"
            "• /removestaff `EMP_ID`\n"
            "• /report — Today's report\n"
            "• /absent — Absent list\n"
            "• /late — Late submissions\n"
            "• /history `EMP_ID` — Last 7 days\n"
            "• /weeklyreport — Weekly grid\n"
            "• /monthly — Monthly stats\n"
            "• /export — Export data\n"
            "• /broadcast `message`\n"
            "• /deadline `HH:MM` — View/set deadline\n"
            "• /sethr `CHAT_ID`\n"
            "\n*Admin (DM):*\n"
            "• /announce `message`\n"
            "• /dm `EMP_ID message`\n"
            "• /remind — Ping pending submitters\n"
            "• /warning `EMP_ID reason`\n"
        )

    await update.message.reply_text(base + admin, parse_mode="Markdown")


# ══════════════════════════════════════════════════════════════════════════════
# /allow — Request re-submission
# ══════════════════════════════════════════════════════════════════════════════

async def allow_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/allow EMP_ID — request approval to re-submit attendance."""
    if not context.args:
        await update.message.reply_text("Usage: /allow EMP_ID")
        return

    emp_id = context.args[0].upper()
    staff = load_staff()
    if emp_id not in staff:
        await update.message.reply_text(f"❌ Unknown employee: {emp_id}")
        return

    att_date = get_attendance_date()
    daily_log = load_daily_log()

    if att_date not in daily_log or emp_id not in daily_log.get(att_date, {}):
        await update.message.reply_text(
            f"ℹ️ {emp_id} has no submission for {att_date}. They can submit directly."
        )
        return

    # Abuse guard — max 1 allow / day / employee
    allow_key = f"{emp_id}:{att_date}"
    count = config.allow_counts.get(allow_key, 0)
    if count >= 1:
        for admin_id in get_admin_ids():
            try:
                await context.bot.send_message(
                    admin_id,
                    f"⚠️ *Suspicious Activity*\n"
                    f"Multiple re-submission requests for {emp_id} "
                    f"({staff[emp_id]['name']}) on {att_date}.",
                    parse_mode="Markdown",
                )
            except Exception:
                pass
        await update.message.reply_text(
            f"⚠️ Re-submission already requested for {emp_id} today. "
            "Additional requests flagged to admins."
        )
        return

    config.allow_counts[allow_key] = count + 1
    config.pending_allows[allow_key] = True

    keyboard = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("✅ Approve", callback_data=f"allow_approve:{emp_id}:{att_date}"),
                InlineKeyboardButton("❌ Reject", callback_data=f"allow_reject:{emp_id}:{att_date}"),
            ]
        ]
    )

    requester = update.effective_user.full_name
    msg = (
        f"🔄 *Re-submission Request*\n\n"
        f"Employee: {staff[emp_id]['name']} ({emp_id})\n"
        f"Date: {att_date}\n"
        f"Requested by: {requester}\n\n"
        f"Approve or reject?"
    )

    for admin_id in get_admin_ids():
        try:
            await context.bot.send_message(
                admin_id, msg, reply_markup=keyboard, parse_mode="Markdown",
            )
        except Exception as exc:
            logger.warning("Failed to send allow request to admin %s: %s", admin_id, exc)

    await update.message.reply_text(
        f"📨 Re-submission request for {emp_id} sent to admin for approval."
    )


# ══════════════════════════════════════════════════════════════════════════════
# Main message handler — attendance submission
# ══════════════════════════════════════════════════════════════════════════════

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Process group text messages in the format ``EMP_ID Work description``.

    * Validates EMP_ID against staff list.
    * Computes attendance date (1 PM cutoff).
    * Enforces one submission per day unless re-submission approved.
    * Persists to daily_log → Excel → Google Sheets.
    * Notifies Owner/HR via DM.
    """
    if not update.message or not update.message.text:
        return
    if update.effective_chat.type not in ("group", "supergroup"):
        return

    # Auto-detect group chat ID
    if config.group_chat_id is None:
        config.group_chat_id = update.effective_chat.id
        logger.info("Group chat ID auto-set to %s", config.group_chat_id)

    text = update.message.text.strip()
    parts = text.split(None, 1)
    if len(parts) < 2:
        return

    emp_id = parts[0].upper()
    work = parts[1].strip()

    staff = load_staff()
    if emp_id not in staff:
        return  # silently ignore unknown IDs

    emp_info = staff[emp_id]
    emp_name = emp_info["name"]
    dept = emp_info.get("dept", "N/A")

    # Link Telegram account on first submission
    user_id = update.effective_user.id
    if not emp_info.get("telegram_id"):
        emp_info["telegram_id"] = user_id
        staff[emp_id] = emp_info
        save_staff(staff)

    att_date = get_attendance_date()
    current_time = now()
    submit_time = current_time.strftime("%H:%M:%S")

    # ── One-submission-per-day check ─────────────────────────────────────
    daily_log = load_daily_log()
    allow_key = f"{emp_id}:{att_date}"
    is_resubmission = False

    if att_date in daily_log and emp_id in daily_log[att_date]:
        if config.approved_resubmissions.get(allow_key):
            is_resubmission = True
            config.approved_resubmissions.pop(allow_key, None)
        else:
            await update.message.reply_text(
                f"⚠️ {emp_name} ({emp_id}), you've already submitted for {att_date}.\n"
                f"Use /allow {emp_id} to request re-submission.",
            )
            return

    # ── Punctuality ──────────────────────────────────────────────────────
    deadline_str = load_deadline()
    try:
        dl_h, dl_m = map(int, deadline_str.split(":"))
    except ValueError:
        dl_h, dl_m = 11, 0

    if current_time.hour < dl_h or (current_time.hour == dl_h and current_time.minute <= dl_m):
        punctuality = "✅ On Time"
        is_late = False
    else:
        punctuality = "❌ LATE"
        is_late = True

    att_dt = datetime.strptime(att_date, "%Y-%m-%d")
    day_name = att_dt.strftime("%A")
    username = update.effective_user.username or update.effective_user.full_name
    source = "Telegram"

    # ── Atomic in-memory update (before any await) ───────────────────────
    daily_log.setdefault(att_date, {})
    daily_log[att_date][emp_id] = {
        "time": submit_time,
        "work": work,
        "late": is_late,
        "username": username,
        "group": update.effective_chat.title or "Group",
        "is_resubmission": is_resubmission,
    }
    save_daily_log(daily_log)

    # ── Persist to Excel (sync) ──────────────────────────────────────────
    excel_handler.save_to_excel(
        emp_id, emp_name, dept, att_date, day_name,
        submit_time, work, source, username, is_resubmission,
    )

    # ── Persist to Google Sheets (async, fire-and-forget) ────────────────
    async def _gs_save_wrapper():
        try:
            await excel_handler.save_to_google_sheets(
                emp_id, emp_name, dept, att_date, day_name,
                submit_time, work, source, username, is_resubmission,
            )
        except Exception as exc:
            logger.error("Google Sheets save failed (background): %s", exc)

    asyncio.create_task(_gs_save_wrapper())

    # ── Confirmation in group ────────────────────────────────────────────
    tag = "🔄 Re-submission" if is_resubmission else "✅ Recorded"
    await update.message.reply_text(
        f"{tag} | {emp_name} ({emp_id}) | {att_date}\n"
        f"⏰ {submit_time}",
    )

    # ── Notify Owner & HR via DM ─────────────────────────────────────────
    resub = " [RE-SUBMISSION]" if is_resubmission else ""
    notify = (
        f"📝 *Attendance Update{resub}*\n\n"
        f"👤 {emp_name} ({emp_id})\n"
        f"🏢 {dept}\n"
        f"📅 {att_date} ({day_name})\n"
        f"⏰ {submit_time}\n"
        f"📋 {work}"
    )
    for admin_id in get_admin_ids():
        try:
            await context.bot.send_message(admin_id, notify, parse_mode="Markdown")
        except Exception as exc:
            logger.warning("Failed to notify admin %s: %s", admin_id, exc)


# ══════════════════════════════════════════════════════════════════════════════
# Post-init — sync data from Google Sheets on cold start
# ══════════════════════════════════════════════════════════════════════════════

async def post_init(application) -> None:
    logger.info("Running post-init data sync …")

    try:
        sheets_data = await excel_handler.load_attendance_from_google_sheets()
        if sheets_data:
            daily_log = load_daily_log()
            merged = 0
            for date_str, entries in sheets_data.items():
                daily_log.setdefault(date_str, {})
                for eid, entry in entries.items():
                    if eid not in daily_log[date_str]:
                        daily_log[date_str][eid] = entry
                        merged += 1
            save_daily_log(daily_log)
            logger.info("Merged %d entries from Google Sheets into daily log.", merged)
    except Exception as exc:
        logger.error("Post-init sync failed: %s", exc)

    # Register slash-command menu
    try:
        await application.bot.set_my_commands(
            [
                BotCommand("start", "Start the bot"),
                BotCommand("help", "Show commands"),
                BotCommand("mystatus", "Weekly status"),
                BotCommand("myprofile", "Your profile"),
                BotCommand("edit", "Request edit"),
                BotCommand("leave", "Request leave"),
                BotCommand("allow", "Request re-submission"),
                BotCommand("staff", "List employees"),
                BotCommand("report", "Today's report"),
                BotCommand("absent", "Absent list"),
                BotCommand("late", "Late list"),
                BotCommand("export", "Export data"),
            ]
        )
    except Exception as exc:
        logger.warning("Failed to set bot commands: %s", exc)

    logger.info("✅ Bot initialized successfully.")


# ══════════════════════════════════════════════════════════════════════════════
# main
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    if not BOT_TOKEN:
        logger.error("BOT_TOKEN not set. Configure your .env file.")
        sys.exit(1)

    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    # Core commands
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("allow", allow_command))

    # Module-level registrations
    register_employee_commands(app)
    register_admin_commands(app)
    register_callbacks(app)

    # Catch-all text handler for attendance (must be added last)
    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND & (filters.ChatType.GROUP | filters.ChatType.SUPERGROUP),
            handle_message,
        )
    )

    logger.info("🚀 SISWIT Attendance Bot starting …")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()
