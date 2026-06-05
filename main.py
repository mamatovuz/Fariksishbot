import json
import logging
import os
import re
import sqlite3
import sys
import time
import warnings
from copy import copy
from datetime import datetime, timezone
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)


load_dotenv(encoding="utf-8-sig")
warnings.filterwarnings("ignore", message=r"If 'per_message=False'.*", category=UserWarning)

DEFAULT_BOT_TOKEN = "8405546234:AAHwIx4gxRxtc-erFp7tckvNLNRlNQjYFV8"
DEFAULT_ADMIN_IDS = "7903688837"
DEFAULT_ADMIN_CHAT_ID = "7903688837"
DEFAULT_PUBLISH_CHAT_ID = "-1003994819171"
DEFAULT_DB_FILE = "bot.db"
DEFAULT_EXCEL_FILE = "applications.xlsx"

BOT_TOKEN = os.getenv("BOT_TOKEN") or DEFAULT_BOT_TOKEN
ADMIN_IDS_TEXT = os.getenv("ADMIN_IDS") or DEFAULT_ADMIN_IDS
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID") or DEFAULT_ADMIN_CHAT_ID
PUBLISH_CHAT_ID = os.getenv("PUBLISH_CHAT_ID") or DEFAULT_PUBLISH_CHAT_ID or ADMIN_CHAT_ID
WEBHOOK_URL = (os.getenv("WEBHOOK_URL") or "").strip()
WEBHOOK_PATH = (os.getenv("WEBHOOK_PATH") or "telegram").strip("/") or "telegram"
RAILWAY_PUBLIC_DOMAIN = os.getenv("RAILWAY_PUBLIC_DOMAIN")
PORT = int(os.getenv("PORT") or "8080")
ADMIN_IDS = {
    int(admin_id.strip())
    for admin_id in ADMIN_IDS_TEXT.split(",")
    if admin_id.strip().lstrip("-").isdigit()
}
DB_FILE = Path(os.getenv("DB_FILE") or DEFAULT_DB_FILE)
EXCEL_FILE = Path(os.getenv("EXCEL_FILE") or DEFAULT_EXCEL_FILE)

(
    FULL_NAME,
    BIRTH_DATE,
    ADDRESS,
    BRANCH,
    EDUCATION,
    EXPERIENCE_CHOICE,
    EXPERIENCE_YEARS,
    PREVIOUS_JOB,
    CONVICTED,
    FAMILY_STATUS,
    PREVIOUS_SALARY,
    EXPECTED_SALARY,
    WORD_LEVEL,
    EXCEL_LEVEL,
    LANGUAGES,
    FARIKS_DURATION,
    MOTIVATION,
    PHONE,
    RECENT_PHOTO,
) = range(19)
ADD_ADMIN_TARGET, REMOVE_ADMIN_TARGET = range(100, 102)

EDUCATION_LABELS = {
    "higher": "Oliy ma'lumotli",
    "secondary": "O'rta maxsus",
}
LEVEL_LABELS = {
    "unknown": "Bilmayman",
    "basic": "Bazaviy",
    "medium": "O'rtacha",
    "good": "Yaxshi",
}
YES_NO_LABELS = {
    "yes": "Ha",
    "no": "Yo'q",
}

EDUCATION_KEYBOARD = InlineKeyboardMarkup(
    [
        [InlineKeyboardButton("🎓 Oliy ma'lumotli", callback_data="education:higher")],
        [InlineKeyboardButton("🏥 O'rta maxsus", callback_data="education:secondary")],
    ]
)
EXPERIENCE_INLINE_KEYBOARD = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("✅ Ha", callback_data="exp:yes"),
            InlineKeyboardButton("❌ Yo'q", callback_data="exp:no"),
        ]
    ]
)
CONVICTED_KEYBOARD = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("✅ Ha", callback_data="convicted:yes"),
            InlineKeyboardButton("❌ Yo'q", callback_data="convicted:no"),
        ]
    ]
)
WORD_LEVEL_KEYBOARD = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("❌ Bilmayman", callback_data="word:unknown"),
            InlineKeyboardButton("🔹 Bazaviy", callback_data="word:basic"),
        ],
        [
            InlineKeyboardButton("🔸 O'rtacha", callback_data="word:medium"),
            InlineKeyboardButton("✅ Yaxshi", callback_data="word:good"),
        ],
    ]
)
EXCEL_LEVEL_KEYBOARD = InlineKeyboardMarkup(
    [
        [
            InlineKeyboardButton("❌ Bilmayman", callback_data="excel:unknown"),
            InlineKeyboardButton("🔹 Bazaviy", callback_data="excel:basic"),
        ],
        [
            InlineKeyboardButton("🔸 O'rtacha", callback_data="excel:medium"),
            InlineKeyboardButton("✅ Yaxshi", callback_data="excel:good"),
        ],
    ]
)
ADMIN_MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [["📝 Ariza yuborish"], ["🛠 Admin panel"]],
    resize_keyboard=True,
)

admin_pending_actions: dict[int, int] = {}
last_conflict_warning_at = 0.0

phone_pattern = re.compile(r"^\+?\d[\d\s()\-]{6,}$")

logging.basicConfig(
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%H:%M:%S",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("telegram").setLevel(logging.WARNING)
logging.getLogger("telegram.ext").setLevel(logging.WARNING)
logging.getLogger("apscheduler").setLevel(logging.WARNING)


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def db_connect() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_FILE)
    connection.row_factory = sqlite3.Row
    return connection


def init_db() -> None:
    with db_connect() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                username TEXT,
                username_lower TEXT,
                first_name TEXT,
                last_name TEXT,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS admins (
                user_id INTEGER PRIMARY KEY,
                username TEXT,
                username_lower TEXT,
                added_by INTEGER,
                added_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS applications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                username TEXT,
                full_name TEXT NOT NULL,
                birth_date TEXT NOT NULL DEFAULT '',
                address TEXT NOT NULL DEFAULT '',
                branch TEXT NOT NULL DEFAULT '',
                education TEXT NOT NULL DEFAULT '',
                phone TEXT NOT NULL,
                experience TEXT NOT NULL,
                previous_job TEXT NOT NULL DEFAULT '',
                convicted TEXT NOT NULL DEFAULT '',
                family_status TEXT NOT NULL DEFAULT '',
                previous_salary TEXT NOT NULL DEFAULT '',
                expected_salary TEXT NOT NULL DEFAULT '',
                word_level TEXT NOT NULL DEFAULT '',
                excel_level TEXT NOT NULL DEFAULT '',
                languages TEXT NOT NULL DEFAULT '',
                fariks_duration TEXT NOT NULL DEFAULT '',
                motivation TEXT NOT NULL DEFAULT '',
                recent_photo_json TEXT NOT NULL DEFAULT '{}',
                role TEXT NOT NULL DEFAULT 'employee',
                age TEXT NOT NULL DEFAULT '',
                direction TEXT NOT NULL DEFAULT '',
                certificates_json TEXT NOT NULL DEFAULT '[]',
                certificate_count INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL,
                created_at TEXT NOT NULL,
                decided_at TEXT,
                decided_by INTEGER
            )
            """
        )
        ensure_columns(connection, "users", {"username_lower": "TEXT"})
        ensure_columns(connection, "admins", {"username_lower": "TEXT"})
        ensure_columns(
            connection,
            "applications",
            {
                "birth_date": "TEXT NOT NULL DEFAULT ''",
                "address": "TEXT NOT NULL DEFAULT ''",
                "branch": "TEXT NOT NULL DEFAULT ''",
                "education": "TEXT NOT NULL DEFAULT ''",
                "previous_job": "TEXT NOT NULL DEFAULT ''",
                "convicted": "TEXT NOT NULL DEFAULT ''",
                "family_status": "TEXT NOT NULL DEFAULT ''",
                "previous_salary": "TEXT NOT NULL DEFAULT ''",
                "expected_salary": "TEXT NOT NULL DEFAULT ''",
                "word_level": "TEXT NOT NULL DEFAULT ''",
                "excel_level": "TEXT NOT NULL DEFAULT ''",
                "fariks_duration": "TEXT NOT NULL DEFAULT ''",
                "motivation": "TEXT NOT NULL DEFAULT ''",
                "recent_photo_json": "TEXT NOT NULL DEFAULT '{}'",
                "role": "TEXT NOT NULL DEFAULT 'employee'",
                "age": "TEXT NOT NULL DEFAULT ''",
                "direction": "TEXT NOT NULL DEFAULT ''",
                "certificates_json": "TEXT NOT NULL DEFAULT '[]'",
                "certificate_count": "INTEGER NOT NULL DEFAULT 0",
            },
        )
        connection.execute("CREATE INDEX IF NOT EXISTS idx_users_username_lower ON users(username_lower)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_admins_username_lower ON admins(username_lower)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_applications_user_id ON applications(user_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_applications_status ON applications(status)")


def ensure_columns(
    connection: sqlite3.Connection,
    table_name: str,
    columns: dict[str, str],
) -> None:
    existing_columns = {
        row["name"] for row in connection.execute(f"PRAGMA table_info({table_name})")
    }
    for column_name, column_type in columns.items():
        if column_name not in existing_columns:
            connection.execute(
                f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}"
            )


def upsert_user(user) -> None:
    if not user:
        return
    with db_connect() as connection:
        connection.execute(
            """
            INSERT INTO users (
                user_id, username, username_lower, first_name, last_name, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                username = excluded.username,
                username_lower = excluded.username_lower,
                first_name = excluded.first_name,
                last_name = excluded.last_name,
                updated_at = excluded.updated_at
            """,
            (
                user.id,
                user.username,
                user.username.lower() if user.username else None,
                user.first_name,
                user.last_name,
                now_iso(),
            ),
        )


def is_admin(user_id: int | None) -> bool:
    if not user_id:
        return False
    if user_id in ADMIN_IDS:
        return True

    with db_connect() as connection:
        row = connection.execute(
            "SELECT 1 FROM admins WHERE user_id = ?",
            (user_id,),
        ).fetchone()
    return row is not None


def get_stats() -> dict:
    with db_connect() as connection:
        users_count = connection.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        db_admin_ids = {
            row["user_id"] for row in connection.execute("SELECT user_id FROM admins")
        }
        pending = connection.execute(
            "SELECT COUNT(*) FROM applications WHERE status = 'pending'"
        ).fetchone()[0]
        approved = connection.execute(
            "SELECT COUNT(*) FROM applications WHERE status = 'approved'"
        ).fetchone()[0]
        rejected = connection.execute(
            "SELECT COUNT(*) FROM applications WHERE status = 'rejected'"
        ).fetchone()[0]

    return {
        "users": users_count,
        "admins": len(db_admin_ids | ADMIN_IDS),
        "pending": pending,
        "approved": approved,
        "rejected": rejected,
    }


def find_user_by_username(username: str) -> dict | None:
    clean_username = username.lstrip("@").lower()
    with db_connect() as connection:
        return connection.execute(
            "SELECT * FROM users WHERE username_lower = ?",
            (clean_username,),
        ).fetchone()


def find_user_by_id(user_id: int) -> dict | None:
    with db_connect() as connection:
        return connection.execute(
            "SELECT * FROM users WHERE user_id = ?",
            (user_id,),
        ).fetchone()


def add_admin(user_id: int, username: str | None, added_by: int) -> None:
    with db_connect() as connection:
        connection.execute(
            """
            INSERT INTO admins (user_id, username, username_lower, added_by, added_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                username = excluded.username,
                username_lower = excluded.username_lower,
                added_by = excluded.added_by,
                added_at = excluded.added_at
            """,
            (
                user_id,
                username,
                username.lower() if username else None,
                added_by,
                now_iso(),
            ),
        )


def remove_admin(user_id: int) -> bool:
    with db_connect() as connection:
        cursor = connection.execute("DELETE FROM admins WHERE user_id = ?", (user_id,))
        return cursor.rowcount > 0


def resolve_admin_target(text: str) -> tuple[int | None, str | None, str | None]:
    target = text.strip()
    if target.lstrip("-").isdigit():
        user_id = int(target)
        known_user = find_user_by_id(user_id)
        username = known_user.get("username") if known_user else None
        return user_id, username, None

    if target.startswith("@"):
        known_user = find_user_by_username(target)
        if known_user:
            return known_user["user_id"], known_user.get("username"), None
        return (
            None,
            None,
            "Bu username bot bazasida topilmadi. Avval u foydalanuvchi botga /start bossin yoki Telegram ID yuboring.",
        )

    return None, None, "Username @ bilan yoki Telegram ID raqam ko'rinishida yuboring."


def get_application(application_id: int) -> sqlite3.Row | None:
    with db_connect() as connection:
        return connection.execute(
            "SELECT * FROM applications WHERE id = ?",
            (application_id,),
        ).fetchone()


def get_all_application_rows() -> list[sqlite3.Row]:
    with db_connect() as connection:
        return list(connection.execute("SELECT * FROM applications ORDER BY id"))


def create_application(user, user_data: dict) -> int:
    photo = user_data.get("recent_photo", {})
    with db_connect() as connection:
        cursor = connection.execute(
            """
            INSERT INTO applications (
                user_id, username, full_name, birth_date, address, branch, education,
                phone, experience, previous_job, convicted, family_status,
                previous_salary, expected_salary, word_level, excel_level, languages,
                fariks_duration, motivation, recent_photo_json, role, age, direction,
                certificates_json, certificate_count, status, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'employee', '', ?, '[]', 0, 'pending', ?)
            """,
            (
                user.id,
                user.username,
                user_data["full_name"],
                user_data["birth_date"],
                user_data["address"],
                user_data["branch"],
                user_data["education"],
                user_data["phone"],
                user_data["experience"],
                user_data["previous_job"],
                user_data["convicted"],
                user_data["family_status"],
                user_data["previous_salary"],
                user_data["expected_salary"],
                user_data["word_level"],
                user_data["excel_level"],
                user_data["languages"],
                user_data["fariks_duration"],
                user_data["motivation"],
                json.dumps(photo, ensure_ascii=False),
                user_data["branch"],
                now_iso(),
            ),
        )
        application_id = cursor.lastrowid

    safe_sync_excel_file()
    return application_id


def delete_application(application_id: int) -> None:
    with db_connect() as connection:
        connection.execute("DELETE FROM applications WHERE id = ?", (application_id,))
    safe_sync_excel_file()


def update_application_status(application_id: int, status: str, admin_id: int) -> None:
    with db_connect() as connection:
        connection.execute(
            """
            UPDATE applications
            SET status = ?, decided_at = ?, decided_by = ?
            WHERE id = ?
            """,
            (status, now_iso(), admin_id, application_id),
        )
    safe_sync_excel_file()


def row_to_application(row: sqlite3.Row | dict) -> dict:
    row_data = dict(row)
    return {
        "id": row_data["id"],
        "user_id": row_data["user_id"],
        "username": row_data.get("username"),
        "full_name": row_data.get("full_name", ""),
        "birth_date": row_data.get("birth_date", ""),
        "address": row_data.get("address", ""),
        "branch": row_data.get("branch", "") or row_data.get("direction", ""),
        "education": row_data.get("education", ""),
        "phone": row_data.get("phone", ""),
        "experience": row_data.get("experience", ""),
        "previous_job": row_data.get("previous_job", ""),
        "convicted": row_data.get("convicted", ""),
        "family_status": row_data.get("family_status", ""),
        "previous_salary": row_data.get("previous_salary", ""),
        "expected_salary": row_data.get("expected_salary", ""),
        "word_level": row_data.get("word_level", ""),
        "excel_level": row_data.get("excel_level", ""),
        "languages": row_data.get("languages", ""),
        "fariks_duration": row_data.get("fariks_duration", ""),
        "motivation": row_data.get("motivation", ""),
        "recent_photo": json.loads(row_data.get("recent_photo_json") or "{}"),
        "status": row_data.get("status", ""),
        "created_at": row_data.get("created_at", ""),
        "decided_at": row_data.get("decided_at", ""),
        "decided_by": row_data.get("decided_by", ""),
    }


def main_keyboard_for(user_id: int) -> ReplyKeyboardMarkup | ReplyKeyboardRemove:
    if is_admin(user_id):
        return ADMIN_MAIN_KEYBOARD
    return ReplyKeyboardRemove()


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    upsert_user(update.effective_user)

    if is_admin(update.effective_user.id):
        await update.message.reply_text(
            "🛠 Admin menyu ochildi.",
            reply_markup=ADMIN_MAIN_KEYBOARD,
        )
        return ConversationHandler.END

    return await begin_application(update, context)


async def begin_application(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    upsert_user(update.effective_user)
    context.user_data.clear()

    await update.message.reply_text(
        "Assalomu alaykum! 👋\n"
        "Fariks jamoasiga ishga ariza topshirish uchun savollarga javob bering.",
        reply_markup=ReplyKeyboardRemove(),
    )
    await update.message.reply_text("1. 👤 Ism-sharifingizni yozing.")
    return FULL_NAME


async def full_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name = update.message.text.strip()
    if len(name) < 3:
        await update.message.reply_text("Iltimos, ism-sharifingizni to'liq yozing.")
        return FULL_NAME

    context.user_data["full_name"] = name
    await update.message.reply_text(
        "2. 🎂 Tug'ilgan kun, oy va yilingizni yozing.\nMasalan: 29.10.2000"
    )
    return BIRTH_DATE


async def birth_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if not is_valid_birth_date(value):
        await update.message.reply_text("Sanani to'g'ri yozing. Masalan: 29.10.2000")
        return BIRTH_DATE

    context.user_data["birth_date"] = value
    await update.message.reply_text("3. 📍 Yashash manzilingizni to'liq yozing.")
    return ADDRESS


def is_valid_birth_date(value: str) -> bool:
    try:
        parsed = datetime.strptime(value, "%d.%m.%Y")
    except ValueError:
        return False
    return datetime(1940, 1, 1) <= parsed <= datetime.now()


async def address(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 5:
        await update.message.reply_text("Iltimos, manzilingizni to'liqroq yozing.")
        return ADDRESS

    context.user_data["address"] = value
    await update.message.reply_text(
        "4. 🏢 Ishlashni xohlagan filialingizni yozing.\nShahar yoki tuman nomini kiriting."
    )
    return BRANCH


async def branch(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 2:
        await update.message.reply_text("Iltimos, filial shahar yoki tuman nomini yozing.")
        return BRANCH

    context.user_data["branch"] = value
    await update.message.reply_text(
        "5. 🎓 Ma'lumot darajangizni tanlang.",
        reply_markup=EDUCATION_KEYBOARD,
    )
    return EDUCATION


async def education_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    key = query.data.split(":", 1)[1]
    context.user_data["education"] = EDUCATION_LABELS[key]
    await query.edit_message_text(f"5. 🎓 Ma'lumot: {EDUCATION_LABELS[key]}")
    await ask_experience(query.message)
    return EXPERIENCE_CHOICE


async def invalid_education(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await reply_to_update(
        update,
        "Iltimos, ma'lumot darajasini tugmalardan tanlang.",
        reply_markup=EDUCATION_KEYBOARD,
    )
    return EDUCATION


async def ask_experience(message) -> None:
    await message.reply_text(
        "6. 💼 Shu sohada umumiy ish tajribangiz bormi?",
        reply_markup=EXPERIENCE_INLINE_KEYBOARD,
    )


async def experience_choice_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    query = update.callback_query
    await query.answer()

    key = query.data.split(":", 1)[1]
    await query.edit_message_reply_markup(reply_markup=None)
    if key == "yes":
        await query.message.reply_text("⏳ Shu sohada umumiy ish tajribangiz necha yil?")
        return EXPERIENCE_YEARS

    context.user_data["experience"] = "0 yil"
    await query.message.reply_text("6. 💼 Tajriba: 0 yil")
    await query.message.reply_text(
        "7. 🏬 Oldingi ish joyingizda necha yil ishlagansiz?\nLavozimingizni ham yozishingiz mumkin."
    )
    return PREVIOUS_JOB


async def invalid_experience_choice(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    await reply_to_update(
        update,
        "Iltimos, tajriba bo'yicha Ha yoki Yo'q tugmasini tanlang.",
        reply_markup=EXPERIENCE_INLINE_KEYBOARD,
    )
    return EXPERIENCE_CHOICE


async def experience_years(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if not value:
        await update.message.reply_text("Iltimos, tajribangizni yozing. Masalan: 3 yil")
        return EXPERIENCE_YEARS

    if value.isdigit():
        value = f"{value} yil"

    context.user_data["experience"] = value
    await update.message.reply_text(
        "7. 🏬 Oldingi ish joyingizda necha yil ishlagansiz?\nLavozimingizni ham yozishingiz mumkin."
    )
    return PREVIOUS_JOB


async def previous_job(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 2:
        await update.message.reply_text("Oldingi ish joyingiz haqida qisqacha yozing.")
        return PREVIOUS_JOB

    context.user_data["previous_job"] = value
    await update.message.reply_text(
        "8. ⚖️ Sudlanganmisiz?",
        reply_markup=CONVICTED_KEYBOARD,
    )
    return CONVICTED


async def convicted_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    key = query.data.split(":", 1)[1]
    context.user_data["convicted"] = YES_NO_LABELS[key]
    await query.edit_message_text(f"8. ⚖️ Sudlangan: {YES_NO_LABELS[key]}")
    await query.message.reply_text(
        "9. 👪 Oilaviy holatingiz qanday? Farzandlaringiz bormi?"
    )
    return FAMILY_STATUS


async def invalid_convicted(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await reply_to_update(
        update,
        "Iltimos, sudlanganlik bo'yicha Ha yoki Yo'q tugmasini tanlang.",
        reply_markup=CONVICTED_KEYBOARD,
    )
    return CONVICTED


async def family_status(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 2:
        await update.message.reply_text("Oilaviy holatingizni qisqacha yozing.")
        return FAMILY_STATUS

    context.user_data["family_status"] = value
    await update.message.reply_text("10. 💰 Oxirgi ish joyingizda qancha maosh olgansiz?")
    return PREVIOUS_SALARY


async def previous_salary(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 2:
        await update.message.reply_text("Iltimos, oldingi maoshingizni yozing.")
        return PREVIOUS_SALARY

    context.user_data["previous_salary"] = value
    await update.message.reply_text("11. 💵 Qancha maoshga ishlashni xohlaysiz?")
    return EXPECTED_SALARY


async def expected_salary(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 2:
        await update.message.reply_text("Iltimos, xohlagan maoshingizni yozing.")
        return EXPECTED_SALARY

    context.user_data["expected_salary"] = value
    await update.message.reply_text(
        "12. 📝 Microsoft Word dasturini bilish darajangizni tanlang.",
        reply_markup=WORD_LEVEL_KEYBOARD,
    )
    return WORD_LEVEL


async def word_level_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    key = query.data.split(":", 1)[1]
    context.user_data["word_level"] = LEVEL_LABELS[key]
    await query.edit_message_text(f"12. 📝 Word: {LEVEL_LABELS[key]}")
    await query.message.reply_text(
        "13. 📊 Microsoft Excel dasturini bilish darajangizni tanlang.",
        reply_markup=EXCEL_LEVEL_KEYBOARD,
    )
    return EXCEL_LEVEL


async def invalid_word_level(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await reply_to_update(
        update,
        "Iltimos, Word darajasini tugmalardan tanlang.",
        reply_markup=WORD_LEVEL_KEYBOARD,
    )
    return WORD_LEVEL


async def excel_level_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    key = query.data.split(":", 1)[1]
    context.user_data["excel_level"] = LEVEL_LABELS[key]
    await query.edit_message_text(f"13. 📊 Excel: {LEVEL_LABELS[key]}")
    await query.message.reply_text(
        "14. 🌐 Qaysi tillarni bilasiz va qay darajada?\nMasalan: O'zbekcha yaxshi, Ruscha o'rtacha"
    )
    return LANGUAGES


async def invalid_excel_level(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await reply_to_update(
        update,
        "Iltimos, Excel darajasini tugmalardan tanlang.",
        reply_markup=EXCEL_LEVEL_KEYBOARD,
    )
    return EXCEL_LEVEL


async def languages(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 2:
        await update.message.reply_text("Iltimos, biladigan tillaringizni yozing.")
        return LANGUAGES

    context.user_data["languages"] = value
    await update.message.reply_text(
        "15. ⏳ “Fariks O'quv markazi”da necha yil ishlash niyatingiz bor?"
    )
    return FARIKS_DURATION


async def fariks_duration(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 1:
        await update.message.reply_text("Iltimos, necha yil ishlash niyatingiz borligini yozing.")
        return FARIKS_DURATION

    context.user_data["fariks_duration"] = value
    await update.message.reply_text(
        "16. ❓ Nima uchun aynan “Fariks”da ishlashni xohlaysiz?"
    )
    return MOTIVATION


async def motivation(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if len(value) < 3:
        await update.message.reply_text("Iltimos, sababini qisqacha yozing.")
        return MOTIVATION

    context.user_data["motivation"] = value
    await update.message.reply_text("17. 📞 Telefon raqamingizni yozing. Masalan: +998901234567")
    return PHONE


async def phone(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    value = update.message.text.strip()
    if not phone_pattern.match(value):
        await update.message.reply_text(
            "Telefon raqamini to'g'ri formatda yuboring. Masalan: +998901234567"
        )
        return PHONE

    context.user_data["phone"] = value
    await update.message.reply_text("18. 📷 Oxirgi 1 oy ichida tushgan rasmingizni yuboring.")
    return RECENT_PHOTO


async def recent_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    image = get_image_info(update)
    if not image:
        await update.message.reply_text("Iltimos, rasm yuboring.")
        return RECENT_PHOTO

    context.user_data["recent_photo"] = image
    return await finish_application(update, context)


def get_image_info(update: Update) -> dict | None:
    message = update.message
    if message.photo:
        return {
            "type": "photo",
            "file_id": message.photo[-1].file_id,
        }

    if message.document and message.document.mime_type:
        if message.document.mime_type.startswith("image/"):
            return {
                "type": "document",
                "file_id": message.document.file_id,
            }

    return None


async def invalid_recent_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Iltimos, oxirgi 1 oy ichida tushgan rasmingizni yuboring.")
    return RECENT_PHOTO


async def finish_application(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user = update.effective_user
    upsert_user(user)

    required_fields = [
        "full_name",
        "birth_date",
        "address",
        "branch",
        "education",
        "experience",
        "previous_job",
        "convicted",
        "family_status",
        "previous_salary",
        "expected_salary",
        "word_level",
        "excel_level",
        "languages",
        "fariks_duration",
        "motivation",
        "phone",
        "recent_photo",
    ]
    if any(not context.user_data.get(field) for field in required_fields):
        await reply_to_update(update, "Ariza ma'lumotlari to'liq emas. /start orqali qayta boshlang.")
        context.user_data.clear()
        return ConversationHandler.END

    application_id = create_application(user, context.user_data)
    application = row_to_application(get_application(application_id))

    try:
        await send_application_to_admin(context, application)
    except Exception as error:
        logger.error("Adminlarga yuborishda xatolik: %s", format_runtime_error(error))
        delete_application(application_id)
        await reply_to_update(
            update,
            "Kechirasiz, arizani adminlarga yuborishda xatolik yuz berdi. "
            "Iltimos, keyinroq qayta urinib ko'ring.",
            reply_markup=main_keyboard_for(user.id),
        )
        context.user_data.clear()
        return ConversationHandler.END

    await reply_to_update(
        update,
        "✅ Arizangiz adminlarga yuborildi. Javobni kuting.",
        reply_markup=main_keyboard_for(user.id),
    )
    context.user_data.clear()
    return ConversationHandler.END


async def reply_to_update(
    update: Update,
    text: str,
    reply_markup=None,
) -> None:
    if update.message:
        await update.message.reply_text(text, reply_markup=reply_markup)
        return

    if update.callback_query and update.callback_query.message:
        await update.callback_query.message.reply_text(text, reply_markup=reply_markup)


async def send_application_to_admin(
    context: ContextTypes.DEFAULT_TYPE,
    application: dict,
) -> None:
    if not ADMIN_CHAT_ID:
        raise RuntimeError("ADMIN_CHAT_ID .env faylida ko'rsatilmagan.")

    await send_application_to_chat(
        context=context,
        chat_id=ADMIN_CHAT_ID,
        application=application,
        reply_markup=build_decision_keyboard(application["id"]),
    )


async def send_application_to_chat(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: str,
    application: dict,
    reply_markup: InlineKeyboardMarkup | None = None,
) -> None:
    caption = build_application_caption(application)
    image = application.get("recent_photo") or {}

    if not image:
        await context.bot.send_message(
            chat_id=chat_id,
            text=limit_telegram_text(caption),
            reply_markup=reply_markup,
        )
        return

    if len(caption) <= 1024:
        await send_single_image(
            context=context,
            chat_id=chat_id,
            image=image,
            caption=caption,
            reply_markup=reply_markup,
        )
        return

    await send_single_image(context=context, chat_id=chat_id, image=image, caption=None)
    await context.bot.send_message(
        chat_id=chat_id,
        text=limit_telegram_text(caption),
        reply_markup=reply_markup,
    )


async def send_single_image(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: str,
    image: dict,
    caption: str | None,
    reply_markup: InlineKeyboardMarkup | None = None,
) -> None:
    if image.get("type") == "photo":
        await context.bot.send_photo(
            chat_id=chat_id,
            photo=image["file_id"],
            caption=caption,
            reply_markup=reply_markup,
        )
        return

    await context.bot.send_document(
        chat_id=chat_id,
        document=image["file_id"],
        caption=caption,
        reply_markup=reply_markup,
    )


def build_application_caption(application: dict) -> str:
    return (
        "🆕 Yangi nomzod\n\n"
        f"👤 Ism-sharif: {application['full_name']}\n"
        f"🎂 Tug'ilgan sana: {application['birth_date']}\n"
        f"📍 Manzil: {application['address']}\n"
        f"🏢 Filial: {application['branch']}\n"
        f"🎓 Ma'lumot: {application['education']}\n"
        f"💼 Soha tajribasi: {application['experience']}\n"
        f"🏬 Oldingi ish joyi: {application['previous_job']}\n"
        f"⚖️ Sudlangan: {application['convicted']}\n"
        f"👪 Oilaviy holati: {application['family_status']}\n"
        f"💰 Oldingi maosh: {application['previous_salary']}\n"
        f"💵 Kutilayotgan maosh: {application['expected_salary']}\n"
        f"📝 Word: {application['word_level']}\n"
        f"📊 Excel: {application['excel_level']}\n"
        f"🌐 Tillar: {application['languages']}\n"
        f"⏳ Fariksda ishlash niyati: {application['fariks_duration']}\n"
        f"❓ Nega Fariks: {application['motivation']}\n"
        f"📞 Telefon: {application['phone']}"
    )


def limit_telegram_text(text: str, limit: int = 3900) -> str:
    if len(text) <= limit:
        return text
    return text[: limit - 20].rstrip() + "\n\n... qisqartirildi"


def build_decision_keyboard(application_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("✅ Tasdiqlash", callback_data=f"app:approve:{application_id}"),
                InlineKeyboardButton("❌ Rad etish", callback_data=f"app:reject:{application_id}"),
            ]
        ]
    )


async def handle_application_decision(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    admin_user = query.from_user
    upsert_user(admin_user)

    if not is_admin(admin_user.id):
        await query.answer("Bu amal faqat adminlar uchun.", show_alert=True)
        return

    _, action, application_id_text = query.data.split(":")
    application_id = int(application_id_text)
    row = get_application(application_id)

    if not row:
        await query.answer("Ariza topilmadi.", show_alert=True)
        return

    application = row_to_application(row)
    if application["status"] != "pending":
        await query.answer("Bu ariza allaqachon ko'rib chiqilgan.", show_alert=True)
        await remove_decision_buttons(query)
        return

    if action == "approve":
        if not PUBLISH_CHAT_ID:
            await query.answer("PUBLISH_CHAT_ID sozlanmagan.", show_alert=True)
            return

        try:
            await send_application_to_chat(
                context=context,
                chat_id=PUBLISH_CHAT_ID,
                application=application,
            )
        except Exception as error:
            logger.error("Kanalga yuborishda xatolik: %s", format_runtime_error(error))
            await query.answer("Kanalga yuborishda xatolik bo'ldi.", show_alert=True)
            return

        update_application_status(application_id, "approved", admin_user.id)
        await notify_user(
            context,
            application["user_id"],
            "Arizangiz tasdiqlandi. Siz bilan yaqinda bog'lanamiz. ✅",
        )
        await remove_decision_buttons(query)
        await query.answer("Ariza tasdiqlandi.")
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"✅ Ariza #{application_id} tasdiqlandi.",
        )
        return

    update_application_status(application_id, "rejected", admin_user.id)
    await notify_user(
        context,
        application["user_id"],
        "Arizangiz rad etildi. ❌",
    )
    await remove_decision_buttons(query)
    await query.answer("Ariza rad etildi.")
    await context.bot.send_message(
        chat_id=query.message.chat_id,
        text=f"❌ Ariza #{application_id} rad etildi.",
    )


async def remove_decision_buttons(query) -> None:
    try:
        await query.edit_message_reply_markup(reply_markup=None)
    except Exception as error:
        logger.warning("Inline tugmalarni o'chirishda xatolik: %s", format_runtime_error(error))


async def notify_user(context: ContextTypes.DEFAULT_TYPE, user_id: int, text: str) -> None:
    try:
        await context.bot.send_message(chat_id=user_id, text=text)
    except Exception as error:
        logger.warning("Foydalanuvchiga xabar bormadi (%s): %s", user_id, format_runtime_error(error))


def sync_excel_file() -> Path:
    EXCEL_FILE.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Arizalar"

    headers = [
        "ID",
        "Sana",
        "Holat",
        "Telegram ID",
        "Username",
        "Ism-sharif",
        "Tug'ilgan sana",
        "Manzil",
        "Filial",
        "Ma'lumot",
        "Soha tajribasi",
        "Oldingi ish joyi",
        "Sudlangan",
        "Oilaviy holati",
        "Oldingi maosh",
        "Kutilayotgan maosh",
        "Word",
        "Excel",
        "Tillar",
        "Fariksda ishlash niyati",
        "Nega Fariks",
        "Telefon",
        "Rasm file_id",
        "Qaror sanasi",
        "Qaror qilgan admin",
    ]
    sheet.append(headers)

    for row in get_all_application_rows():
        application = row_to_application(row)
        photo = application.get("recent_photo") or {}
        sheet.append(
            [
                application["id"],
                format_iso_datetime(application["created_at"]),
                status_label(application["status"]),
                application["user_id"],
                application["username"] or "",
                application["full_name"],
                application["birth_date"],
                application["address"],
                application["branch"],
                application["education"],
                application["experience"],
                application["previous_job"],
                application["convicted"],
                application["family_status"],
                application["previous_salary"],
                application["expected_salary"],
                application["word_level"],
                application["excel_level"],
                application["languages"],
                application["fariks_duration"],
                application["motivation"],
                application["phone"],
                photo.get("file_id", ""),
                format_iso_datetime(application.get("decided_at") or ""),
                application.get("decided_by") or "",
            ]
        )

    for cell in sheet[1]:
        font = copy(cell.font)
        font.bold = True
        cell.font = font

    for column in sheet.columns:
        width = min(
            max(len(str(cell.value or "")) for cell in column) + 2,
            45,
        )
        sheet.column_dimensions[get_column_letter(column[0].column)].width = width

    workbook.save(EXCEL_FILE)
    return EXCEL_FILE


def safe_sync_excel_file() -> Path | None:
    try:
        return sync_excel_file()
    except Exception as error:
        logger.warning("Excel yangilashda xatolik: %s", format_runtime_error(error))
        return None


def status_label(status: str) -> str:
    return {
        "pending": "Kutilmoqda",
        "approved": "Tasdiqlangan",
        "rejected": "Rad etilgan",
    }.get(status, status)


def format_iso_datetime(value: str) -> str:
    if not value:
        return ""
    try:
        return datetime.fromisoformat(value).astimezone().strftime("%d.%m.%Y %H:%M")
    except ValueError:
        return value


async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    upsert_user(update.effective_user)
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Bu bo'lim faqat adminlar uchun.")
        return

    await update.message.reply_text(
        build_admin_panel_text(),
        reply_markup=build_admin_panel_keyboard(),
    )


async def my_id(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    upsert_user(update.effective_user)
    username = f"@{update.effective_user.username}" if update.effective_user.username else "username yo'q"
    await update.message.reply_text(
        f"Sizning Telegram ID: {update.effective_user.id}\nUsername: {username}"
    )


async def admin_panel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    upsert_user(query.from_user)

    if not is_admin(query.from_user.id):
        await query.answer("Bu bo'lim faqat adminlar uchun.", show_alert=True)
        return

    await query.answer()
    await query.edit_message_text(
        build_admin_panel_text(),
        reply_markup=build_admin_panel_keyboard(),
    )


async def admin_excel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    upsert_user(query.from_user)

    if not is_admin(query.from_user.id):
        await query.answer("Bu bo'lim faqat adminlar uchun.", show_alert=True)
        return

    await query.answer("Excel tayyorlanyapti...")
    try:
        excel_path = sync_excel_file()
    except Exception as error:
        logger.error("Excel yuborishda xatolik: %s", format_runtime_error(error))
        await query.message.reply_text(
            "Excel fayl tayyorlashda xatolik bo'ldi. Fayl ochiq bo'lsa yopib qayta urinib ko'ring."
        )
        return

    with excel_path.open("rb") as file:
        await context.bot.send_document(
            chat_id=query.message.chat_id,
            document=file,
            filename=excel_path.name,
            caption="📊 Arizalar Excel fayli",
        )


def build_admin_panel_text() -> str:
    stats = get_stats()
    return (
        "🛠 Admin panel\n\n"
        f"👥 Foydalanuvchilar soni: {stats['users']}\n"
        f"👮 Adminlar soni: {stats['admins']}\n"
        f"📨 Kutilayotgan arizalar: {stats['pending']}\n"
        f"✅ Tasdiqlangan: {stats['approved']}\n"
        f"❌ Rad etilgan: {stats['rejected']}"
    )


def build_admin_panel_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("➕ Admin qo'shish", callback_data="admin:add"),
                InlineKeyboardButton("➖ Admin o'chirish", callback_data="admin:remove"),
            ],
            [InlineKeyboardButton("📊 Excel", callback_data="admin:excel")],
            [InlineKeyboardButton("🔄 Yangilash", callback_data="admin:refresh")],
        ]
    )


async def admin_add_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    upsert_user(query.from_user)

    if not is_admin(query.from_user.id):
        await query.answer("Bu amal faqat adminlar uchun.", show_alert=True)
        return ConversationHandler.END

    await query.answer()
    admin_pending_actions[query.from_user.id] = ADD_ADMIN_TARGET
    await query.message.reply_text(
        "Yangi adminning Telegram ID raqamini yoki @username yuboring.\n"
        "Bekor qilish uchun /cancel.",
    )
    return ConversationHandler.END


async def admin_add_receive(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Bu amal faqat adminlar uchun.")
        admin_pending_actions.pop(update.effective_user.id, None)
        return

    user_id, username, error = resolve_admin_target(update.message.text)
    if error:
        await update.message.reply_text(error)
        return

    add_admin(user_id, username, update.effective_user.id)
    admin_pending_actions.pop(update.effective_user.id, None)
    username_text = f"@{username}" if username else str(user_id)
    await update.message.reply_text(
        f"✅ {username_text} admin qilindi.",
        reply_markup=ADMIN_MAIN_KEYBOARD,
    )


async def admin_remove_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    upsert_user(query.from_user)

    if not is_admin(query.from_user.id):
        await query.answer("Bu amal faqat adminlar uchun.", show_alert=True)
        return

    await query.answer()
    admin_pending_actions[query.from_user.id] = REMOVE_ADMIN_TARGET
    await query.message.reply_text(
        "O'chiriladigan adminning Telegram ID raqamini yoki @username yuboring.\n"
        "Bekor qilish uchun /cancel.",
    )


async def admin_remove_receive(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Bu amal faqat adminlar uchun.")
        admin_pending_actions.pop(update.effective_user.id, None)
        return

    user_id, username, error = resolve_admin_target(update.message.text)
    if error:
        await update.message.reply_text(error)
        return

    if user_id in ADMIN_IDS:
        admin_pending_actions.pop(update.effective_user.id, None)
        await update.message.reply_text(
            "Bu admin .env ichidagi ADMIN_IDS orqali berilgan. Uni paneldan o'chirib bo'lmaydi.",
            reply_markup=ADMIN_MAIN_KEYBOARD,
        )
        return

    removed = remove_admin(user_id)
    admin_pending_actions.pop(update.effective_user.id, None)
    username_text = f"@{username}" if username else str(user_id)
    if removed:
        await update.message.reply_text(
            f"✅ {username_text} adminlikdan olib tashlandi.",
            reply_markup=ADMIN_MAIN_KEYBOARD,
        )
    else:
        await update.message.reply_text(
            "Bu foydalanuvchi adminlar ro'yxatida topilmadi.",
            reply_markup=ADMIN_MAIN_KEYBOARD,
        )


async def admin_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    admin_pending_actions.pop(update.effective_user.id, None)
    await update.message.reply_text("Amal bekor qilindi.", reply_markup=ADMIN_MAIN_KEYBOARD)
    return ConversationHandler.END


async def admin_pending_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    action = admin_pending_actions.get(update.effective_user.id)
    if action == ADD_ADMIN_TARGET:
        await admin_add_receive(update, context)
        return
    if action == REMOVE_ADMIN_TARGET:
        await admin_remove_receive(update, context)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text(
        "Ariza to'ldirish bekor qilindi.",
        reply_markup=main_keyboard_for(update.effective_user.id),
    )
    return ConversationHandler.END


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    global last_conflict_warning_at

    if is_polling_conflict(context.error):
        current_time = time.monotonic()
        if current_time - last_conflict_warning_at > 60:
            logger.warning("Xatolik: %s", format_runtime_error(context.error))
            last_conflict_warning_at = current_time
        return

    logger.error("Xatolik: %s", format_runtime_error(context.error))

    if isinstance(update, Update) and update.effective_message:
        await update.effective_message.reply_text(
            "Kechirasiz, texnik xatolik yuz berdi. Iltimos, qayta urinib ko'ring."
        )


async def on_bot_ready(application: Application) -> None:
    logger.info("Bot ishga tushdi.")


def is_polling_conflict(error: Exception | None) -> bool:
    if error is None:
        return False

    text = str(error).lower()
    return "terminated by other getupdates" in text or (
        "conflict" in text and "getupdates" in text
    )


def get_webhook_base_url() -> str | None:
    if WEBHOOK_URL:
        return WEBHOOK_URL.rstrip("/")
    if RAILWAY_PUBLIC_DOMAIN:
        return f"https://{RAILWAY_PUBLIC_DOMAIN}".rstrip("/")
    return None


def run_bot(application: Application) -> None:
    webhook_base_url = get_webhook_base_url()
    if webhook_base_url:
        webhook_url = f"{webhook_base_url}/{WEBHOOK_PATH}"
        logger.info("Bot webhook rejimida ishga tushmoqda.")
        application.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            url_path=WEBHOOK_PATH,
            webhook_url=webhook_url,
            allowed_updates=Update.ALL_TYPES,
            drop_pending_updates=True,
        )
        return

    logger.info("Bot polling rejimida ishga tushmoqda.")
    application.run_polling(
        allowed_updates=Update.ALL_TYPES,
        drop_pending_updates=True,
    )


def main() -> None:
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN .env faylida ko'rsatilmagan.")
    if not ADMIN_CHAT_ID:
        raise RuntimeError("ADMIN_CHAT_ID .env faylida ko'rsatilmagan.")
    if not ADMIN_IDS:
        logger.warning("ADMIN_IDS bo'sh. Admin panel hech kimga chiqmasligi mumkin.")

    init_db()
    application = Application.builder().token(BOT_TOKEN).post_init(on_bot_ready).build()

    application_conversation = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.Regex(r"^(📝\s*)?Ariza yuborish$"), begin_application),
        ],
        states={
            FULL_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, full_name)],
            BIRTH_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, birth_date)],
            ADDRESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, address)],
            BRANCH: [MessageHandler(filters.TEXT & ~filters.COMMAND, branch)],
            EDUCATION: [
                CallbackQueryHandler(education_choice, pattern=r"^education:(higher|secondary)$"),
                MessageHandler(filters.ALL, invalid_education),
            ],
            EXPERIENCE_CHOICE: [
                CallbackQueryHandler(experience_choice_callback, pattern=r"^exp:(yes|no)$"),
                MessageHandler(filters.ALL, invalid_experience_choice),
            ],
            EXPERIENCE_YEARS: [MessageHandler(filters.TEXT & ~filters.COMMAND, experience_years)],
            PREVIOUS_JOB: [MessageHandler(filters.TEXT & ~filters.COMMAND, previous_job)],
            CONVICTED: [
                CallbackQueryHandler(convicted_choice, pattern=r"^convicted:(yes|no)$"),
                MessageHandler(filters.ALL, invalid_convicted),
            ],
            FAMILY_STATUS: [MessageHandler(filters.TEXT & ~filters.COMMAND, family_status)],
            PREVIOUS_SALARY: [MessageHandler(filters.TEXT & ~filters.COMMAND, previous_salary)],
            EXPECTED_SALARY: [MessageHandler(filters.TEXT & ~filters.COMMAND, expected_salary)],
            WORD_LEVEL: [
                CallbackQueryHandler(word_level_choice, pattern=r"^word:(unknown|basic|medium|good)$"),
                MessageHandler(filters.ALL, invalid_word_level),
            ],
            EXCEL_LEVEL: [
                CallbackQueryHandler(excel_level_choice, pattern=r"^excel:(unknown|basic|medium|good)$"),
                MessageHandler(filters.ALL, invalid_excel_level),
            ],
            LANGUAGES: [MessageHandler(filters.TEXT & ~filters.COMMAND, languages)],
            FARIKS_DURATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, fariks_duration)],
            MOTIVATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, motivation)],
            PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, phone)],
            RECENT_PHOTO: [
                MessageHandler(filters.PHOTO | filters.Document.IMAGE, recent_photo),
                MessageHandler(filters.ALL, invalid_recent_photo),
            ],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CommandHandler("start", start),
        ],
    )

    application.add_handler(application_conversation)
    application.add_handler(
        CallbackQueryHandler(handle_application_decision, pattern=r"^app:(approve|reject):\d+$")
    )
    application.add_handler(CallbackQueryHandler(admin_add_start, pattern=r"^admin:add$"))
    application.add_handler(CallbackQueryHandler(admin_remove_start, pattern=r"^admin:remove$"))
    application.add_handler(CallbackQueryHandler(admin_excel_callback, pattern=r"^admin:excel$"))
    application.add_handler(CallbackQueryHandler(admin_panel_callback, pattern=r"^admin:refresh$"))
    application.add_handler(CommandHandler("myid", my_id))
    application.add_handler(CommandHandler("cancel", admin_cancel))
    application.add_handler(MessageHandler(filters.Regex(r"^(🛠\s*)?Admin panel$"), admin_panel))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, admin_pending_text))
    application.add_error_handler(error_handler)
    try:
        run_bot(application)
    except Exception as error:
        logger.error("Botni ishga tushirishda xatolik: %s", format_runtime_error(error))
        sys.exit(1)


def format_runtime_error(error: Exception | None) -> str:
    if error is None:
        return "Noma'lum xatolik."

    text = str(error)
    lower_text = text.lower()

    if "unauthorized" in lower_text:
        return "Bot token noto'g'ri yoki bot o'chirilgan."
    if "not found" in lower_text and "chat" in lower_text:
        return "Admin yoki kanal ID topilmadi. ADMIN_CHAT_ID/PUBLISH_CHAT_ID ni tekshiring."
    if "forbidden" in lower_text:
        return "Botda ruxsat yo'q. Botni admin guruh/kanalga admin qiling yoki foydalanuvchi botni bloklagan."
    if "terminated by other getupdates" in lower_text or ("conflict" in lower_text and "getupdates" in lower_text):
        return (
            "Bot boshqa joyda ham ishlayapti. Bir xil token bilan faqat bitta bot ishlashi kerak: "
            "local `python main.py`ni to'xtating yoki Railway replicasini 1 ta qiling."
        )
    if "no module named" in lower_text:
        return "Kerakli kutubxona o'rnatilmagan. `pip install -r requirements.txt` qiling."

    return text


if __name__ == "__main__":
    main()
