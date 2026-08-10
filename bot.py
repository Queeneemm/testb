import json
import logging
import os
import re
from datetime import date, datetime, time
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from telegram import Document, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters
from zoneinfo import ZoneInfo

load_dotenv()

logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO)
LOGGER = logging.getLogger(__name__)

DATE_RE = re.compile(r"^(\d{1,2})\.(\d{1,2})$")
HEADERS = ("ФИО", "Дата")
DATA_FILE = Path(os.getenv("DATA_FILE", "data/birthdays.xlsx"))
STATE_FILE = Path(os.getenv("STATE_FILE", "data/sent_notifications.json"))
ADMIN_STATE_FILE = Path(os.getenv("ADMIN_STATE_FILE", "data/admin_ids.json"))
TIMEZONE = ZoneInfo(os.getenv("TIMEZONE", "Europe/Moscow"))
REMINDER_TIME = os.getenv("REMINDER_TIME", "09:00")


def parse_admin_ids() -> set[int]:
    raw = os.getenv("ADMIN_IDS", "")
    return {int(item.strip()) for item in raw.split(",") if item.strip()}


ADMIN_IDS = parse_admin_ids()
ACTION_UPLOAD = "upload"
ACTION_EXPORT = "export"
ACTION_LIST = "list"
ACTION_ADD_ADMIN = "add_admin"
ACTION_REMOVE_ADMIN = "remove_admin"
ACTION_BACK = "back"

MAIN_MENU = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("Загрузить файл", callback_data=ACTION_UPLOAD),
            InlineKeyboardButton("Выгрузить файл", callback_data=ACTION_EXPORT),
        ],
        [InlineKeyboardButton("Просмотреть таблицу", callback_data=ACTION_LIST)],
        [
            InlineKeyboardButton("Добавить админа", callback_data=ACTION_ADD_ADMIN),
            InlineKeyboardButton("Удалить админа", callback_data=ACTION_REMOVE_ADMIN),
        ],
    ]
)
BACK_MENU = InlineKeyboardMarkup([[InlineKeyboardButton("Назад", callback_data=ACTION_BACK)]])
STATE_UPLOAD = "upload"
STATE_ADD_ADMIN = "add_admin"
STATE_REMOVE_ADMIN = "remove_admin"


def load_admin_ids() -> set[int]:
    if ADMIN_STATE_FILE.exists():
        return set(int(item) for item in json.loads(ADMIN_STATE_FILE.read_text(encoding="utf-8")))
    return set(ADMIN_IDS)


def save_admin_ids(admin_ids: set[int]) -> None:
    ADMIN_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    ADMIN_STATE_FILE.write_text(json.dumps(sorted(admin_ids), ensure_ascii=False, indent=2), encoding="utf-8")


def get_admin_ids() -> set[int]:
    return load_admin_ids()


def set_user_state(context: ContextTypes.DEFAULT_TYPE, state: str | None) -> None:
    if state is None:
        context.user_data.pop("state", None)
    else:
        context.user_data["state"] = state


async def show_main_menu(update: Update, text: str = "Главное меню") -> None:
    await update.effective_message.reply_text(text, reply_markup=MAIN_MENU)


async def show_action_menu(update: Update, text: str) -> None:
    await update.effective_message.reply_text(text, reply_markup=BACK_MENU)


def ensure_data_file() -> None:
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    if not DATA_FILE.exists():
        save_birthdays([])


def parse_birthday(value: object) -> str:
    if isinstance(value, datetime):
        value = value.strftime("%d.%m")
    elif isinstance(value, date):
        value = value.strftime("%d.%m")
    else:
        value = str(value or "").strip()

    match = DATE_RE.match(value)
    if not match:
        raise ValueError(f"дата '{value}' должна быть в формате ДД.ММ")

    day = int(match.group(1))
    month = int(match.group(2))
    try:
        date(2000, month, day)
    except ValueError as exc:
        raise ValueError(f"дата '{value}' не существует") from exc
    return f"{day:02d}.{month:02d}"


def read_birthdays(path: Path = DATA_FILE) -> list[dict[str, str]]:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook.active
    if sheet.max_column > 2:
        for row in sheet.iter_rows(min_col=3, values_only=True):
            if any(value is not None for value in row):
                raise ValueError("таблица должна содержать только два столбца: ФИО и Дата")

    headers = [str(sheet.cell(row=1, column=col).value or "").strip() for col in (1, 2)]
    if headers != list(HEADERS):
        raise ValueError("первая строка должна содержать заголовки: ФИО, Дата")

    rows: list[dict[str, str]] = []
    for row_number, row in enumerate(sheet.iter_rows(min_row=2, max_col=2, values_only=True), start=2):
        name_raw, birthday_raw = row
        if name_raw is None and birthday_raw is None:
            continue
        name = str(name_raw or "").strip()
        if not name:
            raise ValueError(f"строка {row_number}: ФИО не может быть пустым")
        rows.append({"name": name, "birthday": parse_birthday(birthday_raw)})
    return rows


def save_birthdays(rows: Iterable[dict[str, str]], path: Path = DATA_FILE) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Birthdays"
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append([row["name"], row["birthday"]])
    sheet.column_dimensions["A"].width = 35
    sheet.column_dimensions["B"].width = 12
    workbook.save(path)


def load_sent_state() -> set[str]:
    if not STATE_FILE.exists():
        return set()
    return set(json.loads(STATE_FILE.read_text(encoding="utf-8")))


def save_sent_state(state: set[str]) -> None:
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(sorted(state), ensure_ascii=False, indent=2), encoding="utf-8")


def is_allowed(update: Update) -> bool:
    user = update.effective_user
    return bool(user and user.id in get_admin_ids())


def restricted(handler):
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        if not is_allowed(update):
            user_id = update.effective_user.id if update.effective_user else "unknown"
            LOGGER.warning("Denied access for Telegram ID %s", user_id)
            if update.effective_message:
                await update.effective_message.reply_text("для доступа обратись к @ioibrieb")
            return
        await handler(update, context)
    return wrapper


MENU_TEXT = "Выберите действие."


@restricted
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    set_user_state(context, None)
    await update.effective_message.reply_text("Кнопки перенесены в сообщения бота.", reply_markup=ReplyKeyboardRemove())
    await show_main_menu(update, MENU_TEXT)


@restricted
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    text = (update.effective_message.text or "").strip()
    state = context.user_data.get("state")

    if text in ("Главное меню", "Назад"):
        set_user_state(context, None)
        await show_main_menu(update)
        return

    if state == STATE_ADD_ADMIN:
        await add_admin(update, context, text)
        return

    if state == STATE_REMOVE_ADMIN:
        await remove_admin(update, context, text)
        return

    if text == "Загрузить файл":
        set_user_state(context, STATE_UPLOAD)
        await show_action_menu(update, "Загрузите .xlsx файл.")
    elif text == "Выгрузить файл":
        await export_table(update, context)
    elif text == "Просмотреть таблицу":
        await list_birthdays(update, context)
    elif text == "Добавить админа":
        set_user_state(context, STATE_ADD_ADMIN)
        await show_action_menu(update, "Введите Telegram ID.")
    elif text == "Удалить админа":
        set_user_state(context, STATE_REMOVE_ADMIN)
        admins = ", ".join(str(item) for item in sorted(get_admin_ids()))
        await show_action_menu(update, f"Введите Telegram ID. Сейчас: {admins}")
    else:
        await show_main_menu(update, MENU_TEXT)


@restricted
async def handle_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()

    action = query.data
    if action == ACTION_BACK:
        set_user_state(context, None)
        await query.edit_message_text("Главное меню", reply_markup=MAIN_MENU)
        return

    if action == ACTION_UPLOAD:
        set_user_state(context, STATE_UPLOAD)
        await query.edit_message_text("Загрузите .xlsx файл.", reply_markup=BACK_MENU)
    elif action == ACTION_EXPORT:
        set_user_state(context, None)
        ensure_data_file()
        await query.edit_message_text("Выгружаю файл.", reply_markup=BACK_MENU)
        await update.effective_message.reply_document(document=DATA_FILE.open("rb"), filename="birthdays.xlsx")
    elif action == ACTION_LIST:
        set_user_state(context, None)
        ensure_data_file()
        rows = read_birthdays()
        text = "Таблица пустая." if not rows else "\n".join(f"{item['name']} — {item['birthday']}" for item in rows)
        await query.edit_message_text(text[:4000], reply_markup=BACK_MENU)
    elif action == ACTION_ADD_ADMIN:
        set_user_state(context, STATE_ADD_ADMIN)
        await query.edit_message_text("Введите Telegram ID.", reply_markup=BACK_MENU)
    elif action == ACTION_REMOVE_ADMIN:
        set_user_state(context, STATE_REMOVE_ADMIN)
        admins = ", ".join(str(item) for item in sorted(get_admin_ids()))
        await query.edit_message_text(f"Введите Telegram ID. Сейчас: {admins}", reply_markup=BACK_MENU)
    else:
        set_user_state(context, None)
        await query.edit_message_text(MENU_TEXT, reply_markup=MAIN_MENU)


async def add_admin(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str) -> None:
    try:
        admin_id = int(text)
    except ValueError:
        await show_action_menu(update, "Нужен Telegram ID числом.")
        return

    admin_ids = get_admin_ids()
    admin_ids.add(admin_id)
    save_admin_ids(admin_ids)
    set_user_state(context, None)
    await show_main_menu(update, f"Админ добавлен: {admin_id}")


async def remove_admin(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str) -> None:
    try:
        admin_id = int(text)
    except ValueError:
        await show_action_menu(update, "Нужен Telegram ID числом.")
        return

    admin_ids = get_admin_ids()
    if admin_id not in admin_ids:
        await show_action_menu(update, f"Админ не найден: {admin_id}")
        return
    if len(admin_ids) == 1:
        await show_action_menu(update, "Нельзя удалить последнего админа.")
        return

    admin_ids.remove(admin_id)
    save_admin_ids(admin_ids)
    set_user_state(context, None)
    await show_main_menu(update, f"Админ удален: {admin_id}")


@restricted
async def export_table(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    ensure_data_file()
    await update.effective_message.reply_document(document=DATA_FILE.open("rb"), filename="birthdays.xlsx")
    await update.effective_message.reply_text("Файл выгружен.", reply_markup=BACK_MENU)


@restricted
async def list_birthdays(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    ensure_data_file()
    rows = read_birthdays()
    if not rows:
        await update.effective_message.reply_text("Таблица пустая.", reply_markup=BACK_MENU)
        return
    text = "\n".join(f"{item['name']} — {item['birthday']}" for item in rows)
    await update.effective_message.reply_text(text[:4000], reply_markup=BACK_MENU)


@restricted
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    state = context.user_data.get("state")
    if state != STATE_UPLOAD:
        await show_main_menu(update, "Нажмите «Загрузить файл» в главном меню.")
        return

    document: Document = update.effective_message.document
    if not document.file_name.lower().endswith(".xlsx"):
        await show_action_menu(update, "Нужен .xlsx файл.")
        return

    telegram_file = await document.get_file()
    with TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir) / "upload.xlsx"
        await telegram_file.download_to_drive(tmp_path)
        try:
            rows = read_birthdays(tmp_path)
        except Exception as exc:
            await show_action_menu(update, f"Ошибка: {exc}")
            return
    save_birthdays(rows)
    set_user_state(context, None)
    await show_main_menu(update, f"Таблица обновлена. Записей: {len(rows)}")


def date_for_year(day: int, month: int, year: int) -> date:
    try:
        return date(year, month, day)
    except ValueError:
        if day == 29 and month == 2:
            return date(year, 3, 1)
        raise


def next_date_for(day_month: str, today: date) -> date:
    day, month = map(int, day_month.split("."))
    candidate = date_for_year(day, month, today.year)
    if candidate < today:
        candidate = date_for_year(day, month, today.year + 1)
    return candidate


async def send_reminders(context: ContextTypes.DEFAULT_TYPE) -> None:
    ensure_data_file()
    today = datetime.now(TIMEZONE).date()
    sent = load_sent_state()
    changed = False

    for item in read_birthdays():
        birthday_date = next_date_for(item["birthday"], today)
        days_left = (birthday_date - today).days
        if days_left not in (0, 10):
            continue

        key = f"{birthday_date.year}:{item['name']}:{item['birthday']}:{days_left}"
        if key in sent:
            continue

        if days_left == 10:
            message = f"У {item['name']} день рождения через 10 дней ({item['birthday']})."
        else:
            message = f"У {item['name']} сегодня день рождения! ({item['birthday']})"

        for chat_id in get_admin_ids():
            await context.bot.send_message(chat_id=chat_id, text=message)
        sent.add(key)
        changed = True

    if changed:
        save_sent_state(sent)


def parse_reminder_time(value: str) -> time:
    hours, minutes = map(int, value.split(":", maxsplit=1))
    return time(hour=hours, minute=minutes, tzinfo=TIMEZONE)


def main() -> None:
    if not os.getenv("BOT_TOKEN"):
        raise RuntimeError("BOT_TOKEN is required")
    if not get_admin_ids():
        raise RuntimeError("ADMIN_IDS is required")
    ensure_data_file()

    application = Application.builder().token(os.environ["BOT_TOKEN"]).build()
    application.add_handler(CommandHandler(["start", "help"], start))
    application.add_handler(CallbackQueryHandler(handle_button))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    application.job_queue.run_daily(send_reminders, time=parse_reminder_time(REMINDER_TIME), name="birthday-reminders")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
