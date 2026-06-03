import json
import logging
import os
import re
import sqlite3
import sys
import warnings
from datetime import datetime, timezone
from pathlib import Path

from dotenv import load_dotenv
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputMediaDocument,
    InputMediaPhoto,
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

BOT_TOKEN = os.getenv("BOT_TOKEN") or DEFAULT_BOT_TOKEN
ADMIN_IDS_TEXT = os.getenv("ADMIN_IDS") or DEFAULT_ADMIN_IDS
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID") or DEFAULT_ADMIN_CHAT_ID
PUBLISH_CHAT_ID = os.getenv("PUBLISH_CHAT_ID") or DEFAULT_PUBLISH_CHAT_ID or ADMIN_CHAT_ID
ADMIN_IDS = {
    int(admin_id.strip())
    for admin_id in ADMIN_IDS_TEXT.split(",")
    if admin_id.strip().lstrip("-").isdigit()
}
DB_FILE = Path(os.getenv("DB_FILE") or DEFAULT_DB_FILE)

FULL_NAME, PHONE, AGE, ROLE_CHOICE, DIRECTION, LANGUAGES, EXPERIENCE_CHOICE, EXPERIENCE_YEARS, CERTIFICATES = range(9)
ADD_ADMIN_TARGET, REMOVE_ADMIN_TARGET = range(100, 102)

ROLE_KEYBOARD = InlineKeyboardMarkup(
    [
        [InlineKeyboardButton("🛠 Admin / Support", callback_data="role:admin")],
        [InlineKeyboardButton("👨‍🏫 O'qituvchi", callback_data="role:teacher")],
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
CERTIFICATE_KEYBOARD = ReplyKeyboardMarkup(
    [["✅ Tugatildi", "⏭ Skip"]],
    resize_keyboard=True,
)
ADMIN_MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [["📝 Ariza yuborish"], ["🛠 Admin panel"]],
    resize_keyboard=True,
)

admin_pending_actions: dict[int, int] = {}

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
                role TEXT NOT NULL DEFAULT 'teacher',
                full_name TEXT NOT NULL,
                phone TEXT NOT NULL,
                age TEXT NOT NULL,
                direction TEXT NOT NULL,
                languages TEXT NOT NULL DEFAULT '',
                experience TEXT NOT NULL,
                certificates_json TEXT NOT NULL,
                certificate_count INTEGER NOT NULL,
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
                "role": "TEXT NOT NULL DEFAULT 'teacher'",
                "languages": "TEXT NOT NULL DEFAULT ''",
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


def get_application(application_id: int) -> dict | None:
    with db_connect() as connection:
        return connection.execute(
            "SELECT * FROM applications WHERE id = ?",
            (application_id,),
        ).fetchone()


def create_application(user, user_data: dict) -> int:
    certificates = user_data.get("certificates", [])
    with db_connect() as connection:
        cursor = connection.execute(
            """
            INSERT INTO applications (
                user_id, username, role, full_name, phone, age, direction, languages,
                experience, certificates_json, certificate_count, status, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?)
            """,
            (
                user.id,
                user.username,
                user_data["role"],
                user_data["full_name"],
                user_data["phone"],
                user_data["age"],
                user_data["direction"],
                user_data.get("languages", ""),
                user_data["experience"],
                json.dumps(certificates, ensure_ascii=False),
                len(certificates),
                now_iso(),
            ),
        )
        return cursor.lastrowid


def delete_application(application_id: int) -> None:
    with db_connect() as connection:
        connection.execute("DELETE FROM applications WHERE id = ?", (application_id,))


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


def row_to_application(row: dict) -> dict:
    row_data = dict(row)
    return {
        "id": row_data["id"],
        "user_id": row_data["user_id"],
        "username": row_data.get("username"),
        "role": row_data.get("role", "teacher"),
        "full_name": row_data["full_name"],
        "phone": row_data["phone"],
        "age": row_data["age"],
        "direction": row_data.get("direction", ""),
        "languages": row_data.get("languages", ""),
        "experience": row_data["experience"],
        "certificates": json.loads(row_data.get("certificates_json") or "[]"),
        "certificate_count": row_data.get("certificate_count", 0),
        "status": row_data["status"],
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
    context.user_data["certificates"] = []

    await update.message.reply_text(
        "Assalomu alaykum! 👋\n"
        "Fariks Aloqa jamoasiga ariza yuborish uchun savollarga javob bering.",
        reply_markup=ReplyKeyboardRemove(),
    )
    await update.message.reply_text("👤 Ism va familiyangizni kiriting.")
    return FULL_NAME


async def full_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    name = update.message.text.strip()
    if len(name) < 3:
        await update.message.reply_text("Iltimos, ism va familiyangizni to'liq kiriting.")
        return FULL_NAME

    context.user_data["full_name"] = name
    await update.message.reply_text("📞 Telefon raqamingizni yuboring. Masalan: +998901234567")
    return PHONE


async def phone(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    phone_number = update.message.text.strip()
    if not phone_pattern.match(phone_number):
        await update.message.reply_text(
            "Telefon raqamini to'g'ri formatda yuboring. Masalan: +998901234567"
        )
        return PHONE

    context.user_data["phone"] = phone_number
    await update.message.reply_text("🎂 Yoshingiz nechida?")
    return AGE


async def age(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    age_text = update.message.text.strip()
    if not age_text.isdigit() or not 1 <= int(age_text) <= 120:
        await update.message.reply_text("Iltimos, yoshingizni raqam bilan kiriting.")
        return AGE

    context.user_data["age"] = age_text
    await update.message.reply_text(
        "🚩 Qaysi yo'nalish bo'yicha ariza topshirmoqchisiz?",
        reply_markup=ROLE_KEYBOARD,
    )
    return ROLE_CHOICE


async def role_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    role = query.data.split(":", 1)[1]
    context.user_data["role"] = role
    await query.edit_message_reply_markup(reply_markup=None)

    if role == "teacher":
        context.user_data["direction_type"] = "teacher"
        await query.message.reply_text(
            "📚 Qaysi fan yoki yo'nalishni o'qitasiz? Masalan: Matematika"
        )
        return DIRECTION

    context.user_data["direction"] = "Admin / Support"
    context.user_data["certificates"] = []
    await query.message.reply_text("🌐 Qaysi tillarni bilasiz? Masalan: O'zbek, Rus, Ingliz")
    return LANGUAGES


async def invalid_role_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await reply_to_update(
        update,
        "Iltimos, quyidagi tugmalardan birini tanlang.",
        reply_markup=ROLE_KEYBOARD,
    )
    return ROLE_CHOICE


async def languages(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    languages_text = update.message.text.strip()
    if len(languages_text) < 2:
        await update.message.reply_text("Iltimos, biladigan tillaringizni yozing.")
        return LANGUAGES

    context.user_data["languages"] = languages_text
    await ask_experience(update)
    return EXPERIENCE_CHOICE


async def invalid_languages(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await reply_to_update(update, "🌐 Iltimos, tillarni matn ko'rinishida yozing.")
    return LANGUAGES


async def ask_experience(update: Update) -> None:
    message = update.callback_query.message if update.callback_query else update.message
    await message.reply_text(
        "💼 Ushbu sohada tajribangiz bormi?",
        reply_markup=EXPERIENCE_INLINE_KEYBOARD,
    )


async def finish_after_experience(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    if context.user_data.get("role") == "admin":
        return await certificates_done(update, context)

    await ask_certificates(update)
    return CERTIFICATES


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


async def experience_choice_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    query = update.callback_query
    await query.answer()
    await query.edit_message_reply_markup(reply_markup=None)

    choice = query.data.split(":", 1)[1]
    if choice == "yes":
        await query.message.reply_text(
            "⏳ Necha yil yoki oy tajribangiz bor? Masalan: 2 yil yoki 6 oy",
            reply_markup=ReplyKeyboardRemove(),
        )
        return EXPERIENCE_YEARS

    context.user_data["experience"] = "0 yil"
    return await finish_after_experience(update, context)


async def invalid_experience_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    await ask_experience(update)
    return EXPERIENCE_CHOICE


async def direction(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    direction_text = update.message.text.strip()
    if len(direction_text) < 2:
        await update.message.reply_text("Iltimos, fan yoki yo'nalishni kiriting.")
        return DIRECTION

    context.user_data["direction"] = direction_text
    context.user_data["languages"] = ""
    await ask_experience(update)
    return EXPERIENCE_CHOICE


async def invalid_direction(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await reply_to_update(update, "📚 Iltimos, fan yoki yo'nalishni matn ko'rinishida yozing.")
    return DIRECTION


async def experience_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    answer = update.message.text.strip().lower()

    if "ha" in answer:
        await update.message.reply_text(
            "⏳ Necha yil yoki oy tajribangiz bor? Masalan: 2 yil yoki 6 oy",
            reply_markup=ReplyKeyboardRemove(),
        )
        return EXPERIENCE_YEARS

    if "yo'q" in answer or "yoq" in answer:
        context.user_data["experience"] = "0 yil"
        return await finish_after_experience(update, context)

    await ask_experience(update)
    return EXPERIENCE_CHOICE


async def experience_years(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    experience = update.message.text.strip()
    if not experience:
        await update.message.reply_text("Iltimos, tajribangizni kiriting. Masalan: 3 yil")
        return EXPERIENCE_YEARS

    if experience.isdigit():
        experience = f"{experience} yil"

    context.user_data["experience"] = experience
    return await finish_after_experience(update, context)


async def ask_certificates(update: Update) -> None:
    await update.message.reply_text(
        "📜 Sertifikatlaringiz rasmlarini yuboring. Bir nechta sertifikat bo'lsa, "
        "barchasini yuborishingiz mumkin.\n\n"
        "Sertifikat bo'lmasa \"⏭ Skip\" tugmasini bosing yoki /skip yuboring.\n"
        "Barcha rasmlarni yuborib bo'lgach, \"✅ Tugatildi\" tugmasini bosing.",
        reply_markup=CERTIFICATE_KEYBOARD,
    )


async def save_certificate(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    certificate = get_certificate_info(update)
    if not certificate:
        await update.message.reply_text(
            "📜 Iltimos, sertifikat rasmini yuboring, \"✅ Tugatildi\" yoki \"⏭ Skip\" tugmasini bosing.",
            reply_markup=CERTIFICATE_KEYBOARD,
        )
        return CERTIFICATES

    context.user_data.setdefault("certificates", []).append(certificate)

    await update.message.reply_text(
        "✅ Sertifikat qabul qilindi. Yana rasm yuborishingiz yoki "
        "\"✅ Tugatildi\" tugmasini bosishingiz mumkin.",
        reply_markup=CERTIFICATE_KEYBOARD,
    )
    return CERTIFICATES


def get_certificate_info(update: Update) -> dict | None:
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


async def certificates_done(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user = update.effective_user
    upsert_user(user)

    required_fields = ["full_name", "phone", "age", "role", "direction", "experience"]
    if context.user_data.get("role") == "admin":
        required_fields.append("languages")

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


async def send_application_to_admin(
    context: ContextTypes.DEFAULT_TYPE,
    application: dict,
) -> None:
    if not ADMIN_CHAT_ID:
        raise RuntimeError("ADMIN_CHAT_ID .env faylida ko'rsatilmagan.")

    keyboard = build_decision_keyboard(application["id"])
    await send_application_to_chat(
        context=context,
        chat_id=ADMIN_CHAT_ID,
        application=application,
        reply_markup=keyboard,
        controls_for_album=True,
    )


async def send_application_to_chat(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: str,
    application: dict,
    reply_markup: InlineKeyboardMarkup | None = None,
    controls_for_album: bool = False,
) -> None:
    caption = build_application_caption(application)
    certificates = application["certificates"]

    if not certificates:
        await context.bot.send_message(chat_id=chat_id, text=caption, reply_markup=reply_markup)
        return

    if len(certificates) == 1:
        await send_single_certificate(
            context=context,
            chat_id=chat_id,
            certificate=certificates[0],
            caption=caption,
            reply_markup=reply_markup,
        )
        return

    await send_certificates_album(context, chat_id, certificates, caption)

    if reply_markup and controls_for_album:
        await context.bot.send_message(
            chat_id=chat_id,
            text=f"Ariza #{application['id']} bo'yicha qaror tanlang:",
            reply_markup=reply_markup,
        )


async def send_single_certificate(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: str,
    certificate: dict,
    caption: str | None,
    reply_markup: InlineKeyboardMarkup | None = None,
) -> None:
    if certificate["type"] == "photo":
        await context.bot.send_photo(
            chat_id=chat_id,
            photo=certificate["file_id"],
            caption=caption,
            reply_markup=reply_markup,
        )
        return

    await context.bot.send_document(
        chat_id=chat_id,
        document=certificate["file_id"],
        caption=caption,
        reply_markup=reply_markup,
    )


async def send_certificates_album(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: str,
    certificates: list[dict],
    caption: str,
) -> None:
    caption_sent = False
    for certificates_chunk in build_certificate_chunks(certificates):
        chunk_caption = caption if not caption_sent else None
        if len(certificates_chunk) == 1:
            await send_single_certificate(
                context=context,
                chat_id=chat_id,
                certificate=certificates_chunk[0],
                caption=chunk_caption,
            )
        else:
            await context.bot.send_media_group(
                chat_id=chat_id,
                media=build_media_group(certificates_chunk, chunk_caption),
            )
        caption_sent = True


def build_certificate_chunks(certificates: list[dict]) -> list[list[dict]]:
    chunks = []
    current_chunk = []
    current_type = None

    for certificate in certificates:
        if (
            current_chunk
            and (certificate["type"] != current_type or len(current_chunk) == 10)
        ):
            chunks.append(current_chunk)
            current_chunk = []

        current_chunk.append(certificate)
        current_type = certificate["type"]

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


def build_media_group(certificates: list[dict], caption: str | None) -> list:
    media_group = []

    for index, certificate in enumerate(certificates):
        item_caption = caption if index == 0 else None
        if certificate["type"] == "photo":
            media_group.append(InputMediaPhoto(media=certificate["file_id"], caption=item_caption))
        else:
            media_group.append(InputMediaDocument(media=certificate["file_id"], caption=item_caption))

    return media_group


def build_application_caption(application: dict) -> str:
    text = (
        "🆕 Yangi nomzod\n\n"
        f"👤 Ism: {application['full_name']}\n"
        f"📞 Telefon: {application['phone']}\n"
        f"🎂 Yosh: {application['age']}\n"
        f"💼 Tajriba: {application['experience']}\n"
        f"🚩 Yo'nalish: {application['direction']}\n"
    )
    if application.get("languages"):
        text += f"🌐 Tillar: {application['languages']}\n"
    if application.get("role") != "admin":
        certificates_text = (
            f"{application['certificate_count']} ta"
            if application["certificate_count"]
            else "Sertifikat yo'q"
        )
        text += f"📜 Sertifikatlar: {certificates_text}"
    else:
        text = text.rstrip()
    return text


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
        "Arizangiz rad etildi. Yangi ariza yuborishingiz mumkin.",
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
            [InlineKeyboardButton("🔄 Yangilash", callback_data="admin:refresh")],
        ]
    )


async def admin_add_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    upsert_user(query.from_user)

    if not is_admin(query.from_user.id):
        await query.answer("Bu amal faqat adminlar uchun.", show_alert=True)
        return

    await query.answer()
    admin_pending_actions[query.from_user.id] = ADD_ADMIN_TARGET
    await query.message.reply_text(
        "Yangi adminning Telegram ID raqamini yoki @username yuboring.\n"
        "Bekor qilish uchun /cancel.",
    )


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
        return


async def invalid_experience_choice(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    await ask_experience(update)
    return EXPERIENCE_CHOICE


async def invalid_certificate(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    await update.message.reply_text(
        "Iltimos, sertifikat rasmini yuboring, \"Tugatildi\" yoki \"Skip\" tugmasini bosing.",
        reply_markup=CERTIFICATE_KEYBOARD,
    )
    return CERTIFICATES


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text(
        "Ariza to'ldirish bekor qilindi.",
        reply_markup=main_keyboard_for(update.effective_user.id),
    )
    return ConversationHandler.END


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.error("Xatolik: %s", format_runtime_error(context.error))

    if isinstance(update, Update) and update.effective_message:
        await update.effective_message.reply_text(
            "Kechirasiz, texnik xatolik yuz berdi. Iltimos, qayta urinib ko'ring."
        )


async def on_bot_ready(application: Application) -> None:
    logger.info("Bot ishga tushdi.")


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
            FULL_NAME: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, full_name),
            ],
            PHONE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, phone),
            ],
            AGE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, age),
            ],
            ROLE_CHOICE: [
                CallbackQueryHandler(role_choice, pattern=r"^role:(admin|teacher)$"),
                MessageHandler(filters.ALL, invalid_role_choice),
            ],
            DIRECTION: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, direction),
                MessageHandler(filters.ALL, invalid_direction),
            ],
            LANGUAGES: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, languages),
                MessageHandler(filters.ALL, invalid_languages),
            ],
            EXPERIENCE_CHOICE: [
                CallbackQueryHandler(experience_choice_callback, pattern=r"^exp:(yes|no)$"),
                MessageHandler(
                    filters.Regex(r"^(✅\s*)?Ha$|^(❌\s*)?Yo'?q$"),
                    experience_choice,
                ),
                MessageHandler(filters.TEXT & ~filters.COMMAND, invalid_experience_choice),
                MessageHandler(filters.ALL, invalid_experience_callback),
            ],
            EXPERIENCE_YEARS: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, experience_years),
            ],
            CERTIFICATES: [
                CommandHandler("skip", certificates_done),
                MessageHandler(
                    filters.Regex(r"(?i)^((✅\s*)?Tugatildi|(⏭\s*)?skip)$"),
                    certificates_done,
                ),
                MessageHandler(filters.PHOTO | filters.Document.IMAGE, save_certificate),
                MessageHandler(filters.ALL, invalid_certificate),
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
    application.add_handler(CallbackQueryHandler(admin_panel_callback, pattern=r"^admin:refresh$"))
    application.add_handler(CommandHandler("myid", my_id))
    application.add_handler(CommandHandler("cancel", admin_cancel))
    application.add_handler(MessageHandler(filters.Regex(r"^(🛠\s*)?Admin panel$"), admin_panel))
    application.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, admin_pending_text)
    )
    application.add_error_handler(error_handler)
    try:
        application.run_polling(allowed_updates=Update.ALL_TYPES)
    except Exception as error:
        logger.error("Botni ishga tushirishda xatolik: %s", format_runtime_error(error))
        sys.exit(1)


def format_runtime_error(error: Exception | None) -> str:
    if error is None:
        return "Noma'lum xatolik."

    text = str(error)
    lowered = text.lower()

    if "unauthorized" in lowered:
        return "Bot token xato yoki bekor qilingan. Tokenni tekshiring."
    if "chat not found" in lowered:
        return "Admin yoki kanal ID topilmadi. ADMIN_CHAT_ID/PUBLISH_CHAT_ID ni tekshiring."
    if "forbidden" in lowered:
        return "Botda ruxsat yo'q. Bot kanal/guruhda adminmi yoki user botni start qilganmi, tekshiring."
    if "database is locked" in lowered:
        return "SQLite bazasi band. Botni ikki marta ishga tushirmaganingizni tekshiring."
    if not text:
        return error.__class__.__name__
    return text


if __name__ == "__main__":
    main()
