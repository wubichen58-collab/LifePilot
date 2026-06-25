import asyncio
import base64
import logging
import os
import re
import sqlite3
import tempfile
from datetime import datetime, timedelta
from io import BytesIO
from typing import Any, Dict, List, Optional
from urllib.parse import quote_plus

import aiohttp
import edge_tts
import openpyxl
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandObject
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from aiohttp import web
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from docx import Document as DocxDocument
from pypdf import PdfReader


# =========================
# CONFIG
# =========================

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
QWEATHER_KEY = os.environ.get("QWEATHER_KEY")
JARVIS_API_KEY = os.environ.get("JARVIS_API_KEY")  # секрет для /jarvis HTTP endpoint (MacroDroid)
JARVIS_HTTP_USER_ID = int(os.environ.get("JARVIS_HTTP_USER_ID", "0"))  # твой Telegram user_id, чтобы шарить память/задачи

GROQ_CHAT_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_TRANSCRIBE_URL = "https://api.groq.com/openai/v1/audio/transcriptions"

GROQ_MODEL_TEXT = "llama-3.3-70b-versatile"
GROQ_MODEL_FAST = "llama-3.1-8b-instant"
GROQ_MODEL_VISION = "meta-llama/llama-4-scout-17b-16e-instruct"
GROQ_MODEL_ASR = "whisper-large-v3"

VOICE_NAME = "ru-RU-DmitryNeural"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "jarvis.db")

os.makedirs(DATA_DIR, exist_ok=True)

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is required")


# =========================
# LOGGING
# =========================

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("jarvis")


# =========================
# BOT / APP
# =========================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
router = Router()
dp.include_router(router)

scheduler = AsyncIOScheduler()
DB_LOCK = asyncio.Lock()
HTTP_SESSION: Optional[aiohttp.ClientSession] = None


# =========================
# DB
# =========================

def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                user_id INTEGER PRIMARY KEY,
                voice_mode INTEGER DEFAULT 0,
                model_mode TEXT DEFAULT 'smart',
                briefing_enabled INTEGER DEFAULT 0,
                briefing_time TEXT DEFAULT '09:00',
                default_city TEXT DEFAULT NULL,
                last_brief_date TEXT DEFAULT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                role TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS memories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                text TEXT NOT NULL,
                status TEXT DEFAULT 'open',
                due_at TEXT DEFAULT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS reminders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                chat_id INTEGER NOT NULL,
                text TEXT NOT NULL,
                due_at TEXT NOT NULL,
                sent INTEGER DEFAULT 0,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        conn.commit()


async def db_execute(sql: str, params: tuple = ()) -> int:
    async with DB_LOCK:
        def _run() -> int:
            with sqlite3.connect(DB_PATH) as conn:
                conn.row_factory = sqlite3.Row
                cur = conn.execute(sql, params)
                conn.commit()
                return int(cur.lastrowid or 0)

        return await asyncio.to_thread(_run)


async def db_query(sql: str, params: tuple = ()) -> List[Dict[str, Any]]:
    async with DB_LOCK:
        def _run() -> List[Dict[str, Any]]:
            with sqlite3.connect(DB_PATH) as conn:
                conn.row_factory = sqlite3.Row
                cur = conn.execute(sql, params)
                rows = cur.fetchall()
                return [dict(r) for r in rows]

        return await asyncio.to_thread(_run)


async def ensure_user(user_id: int) -> None:
    await db_execute("INSERT OR IGNORE INTO settings (user_id) VALUES (?)", (user_id,))


async def get_settings(user_id: int) -> Dict[str, Any]:
    await ensure_user(user_id)
    rows = await db_query("SELECT * FROM settings WHERE user_id = ?", (user_id,))
    return rows[0] if rows else {
        "user_id": user_id,
        "voice_mode": 0,
        "model_mode": "smart",
        "briefing_enabled": 0,
        "briefing_time": "09:00",
        "default_city": None,
        "last_brief_date": None,
    }


async def set_setting(user_id: int, key: str, value: Any) -> None:
    await ensure_user(user_id)
    await db_execute(f"UPDATE settings SET {key} = ? WHERE user_id = ?", (value, user_id))


async def save_message(user_id: int, role: str, content: str) -> None:
    await db_execute(
        "INSERT INTO messages (user_id, role, content) VALUES (?, ?, ?)",
        (user_id, role, content),
    )


async def get_recent_messages(user_id: int, limit: int = 12) -> List[Dict[str, str]]:
    rows = await db_query(
        """
        SELECT role, content
        FROM messages
        WHERE user_id = ?
        ORDER BY id DESC
        LIMIT ?
        """,
        (user_id, limit),
    )
    rows.reverse()
    return [{"role": r["role"], "content": r["content"]} for r in rows]


async def clear_history(user_id: int) -> None:
    await db_execute("DELETE FROM messages WHERE user_id = ?", (user_id,))


async def add_memory(user_id: int, content: str) -> int:
    return await db_execute(
        "INSERT INTO memories (user_id, content) VALUES (?, ?)",
        (user_id, content),
    )


async def get_memories(user_id: int, limit: int = 20) -> List[Dict[str, Any]]:
    return await db_query(
        """
        SELECT id, content, created_at
        FROM memories
        WHERE user_id = ?
        ORDER BY id DESC
        LIMIT ?
        """,
        (user_id, limit),
    )


async def delete_memory(user_id: int, memory_id: int) -> None:
    await db_execute(
        "DELETE FROM memories WHERE user_id = ? AND id = ?",
        (user_id, memory_id),
    )


async def add_note(user_id: int, title: str, content: str) -> int:
    return await db_execute(
        "INSERT INTO notes (user_id, title, content) VALUES (?, ?, ?)",
        (user_id, title, content),
    )


async def get_notes(user_id: int, limit: int = 20) -> List[Dict[str, Any]]:
    return await db_query(
        """
        SELECT id, title, content, created_at
        FROM notes
        WHERE user_id = ?
        ORDER BY id DESC
        LIMIT ?
        """,
        (user_id, limit),
    )


async def add_task(user_id: int, text: str, due_at: Optional[str] = None) -> int:
    return await db_execute(
        "INSERT INTO tasks (user_id, text, due_at) VALUES (?, ?, ?)",
        (user_id, text, due_at),
    )


async def get_tasks(user_id: int, status: str = "open", limit: int = 20) -> List[Dict[str, Any]]:
    return await db_query(
        """
        SELECT id, text, status, due_at, created_at
        FROM tasks
        WHERE user_id = ? AND status = ?
        ORDER BY id DESC
        LIMIT ?
        """,
        (user_id, status, limit),
    )


async def mark_task_done(user_id: int, task_id: int) -> bool:
    affected = await db_execute(
        "UPDATE tasks SET status = 'done' WHERE user_id = ? AND id = ?",
        (user_id, task_id),
    )
    return affected > 0


async def add_reminder(user_id: int, chat_id: int, text: str, due_at: str) -> int:
    return await db_execute(
        """
        INSERT INTO reminders (user_id, chat_id, text, due_at)
        VALUES (?, ?, ?, ?)
        """,
        (user_id, chat_id, text, due_at),
    )


async def get_due_reminders(limit: int = 100) -> List[Dict[str, Any]]:
    now_iso = datetime.now().isoformat(timespec="minutes")
    return await db_query(
        """
        SELECT id, user_id, chat_id, text, due_at
        FROM reminders
        WHERE sent = 0 AND due_at <= ?
        ORDER BY due_at ASC
        LIMIT ?
        """,
        (now_iso, limit),
    )


async def mark_reminder_sent(reminder_id: int) -> None:
    await db_execute(
        "UPDATE reminders SET sent = 1 WHERE id = ?",
        (reminder_id,),
    )


async def get_due_reminders_for_user(user_id: int) -> List[Dict[str, Any]]:
    now_iso = datetime.now().isoformat(timespec="minutes")
    return await db_query(
        """
        SELECT id, user_id, chat_id, text, due_at
        FROM reminders
        WHERE sent = 0 AND user_id = ? AND due_at <= ?
        ORDER BY due_at ASC
        """,
        (user_id, now_iso),
    )


# =========================
# UTILITIES
# =========================

def split_text(text: str, limit: int = 3500) -> List[str]:
    text = text.strip()
    if len(text) <= limit:
        return [text]

    parts = []
    current = ""
    for paragraph in text.split("\n"):
        candidate = current + ("\n" if current else "") + paragraph
        if len(candidate) <= limit:
            current = candidate
            continue
        if current:
            parts.append(current)
            current = ""

        words = paragraph.split()
        buffer = ""
        for word in words:
            cand = buffer + (" " if buffer else "") + word
            if len(cand) > limit:
                if buffer:
                    parts.append(buffer)
                buffer = word
            else:
                buffer = cand
        if buffer:
            parts.append(buffer)

    if current:
        parts.append(current)

    return [p for p in parts if p.strip()]


def clip(text: str, n: int = 12000) -> str:
    return text if len(text) <= n else text[:n]


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def build_main_keyboard(settings: Dict[str, Any]) -> InlineKeyboardMarkup:
    voice = "🔊 Voice: ON" if int(settings.get("voice_mode", 0)) else "🔇 Voice: OFF"
    brief = "🗓 Briefing: ON" if int(settings.get("briefing_enabled", 0)) else "🗓 Briefing: OFF"

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🧠 Память", callback_data="jarvis_memory"),
                InlineKeyboardButton(text="📌 Задачи", callback_data="jarvis_tasks"),
            ],
            [
                InlineKeyboardButton(text="⏰ Напоминание", callback_data="jarvis_remind"),
                InlineKeyboardButton(text="🌦 Погода", callback_data="jarvis_weather"),
            ],
            [
                InlineKeyboardButton(text=voice, callback_data="jarvis_voice"),
                InlineKeyboardButton(text=brief, callback_data="jarvis_briefing"),
            ],
        ]
    )


def select_model(settings: Dict[str, Any]) -> str:
    mode = str(settings.get("model_mode", "smart")).lower()
    if mode == "fast":
        return GROQ_MODEL_FAST
    if mode == "vision":
        return GROQ_MODEL_VISION
    return GROQ_MODEL_TEXT


# =========================
# HTTP SESSION
# =========================

async def get_http_session() -> aiohttp.ClientSession:
    global HTTP_SESSION
    if HTTP_SESSION is None or HTTP_SESSION.closed:
        timeout = aiohttp.ClientTimeout(total=120)
        HTTP_SESSION = aiohttp.ClientSession(timeout=timeout)
    return HTTP_SESSION


# =========================
# GROQ
# =========================

async def groq_chat(messages: List[Dict[str, Any]], model: str, temperature: float = 0.5) -> str:
    if not GROQ_API_KEY:
        return "GROQ_API_KEY не задан."

    session = await get_http_session()
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
    }

    async with session.post(GROQ_CHAT_URL, headers=headers, json=payload) as resp:
        data = await resp.json(content_type=None)
        if resp.status >= 400:
            logger.error("Groq error %s: %s", resp.status, data)
            return f"Ошибка Groq API: {data}"
        try:
            return data["choices"][0]["message"]["content"].strip()
        except Exception:
            logger.error("Unexpected Groq response: %s", data)
            return "Не удалось разобрать ответ модели."


async def groq_transcribe(audio_bytes: bytes, filename: str = "voice.ogg") -> Optional[str]:
    if not GROQ_API_KEY:
        return None

    session = await get_http_session()
    form = aiohttp.FormData()
    form.add_field(
        "file",
        audio_bytes,
        filename=filename,
        content_type="audio/ogg",
    )
    form.add_field("model", GROQ_MODEL_ASR)

    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}

    async with session.post(GROQ_TRANSCRIBE_URL, headers=headers, data=form) as resp:
        data = await resp.json(content_type=None)
        if resp.status >= 400:
            logger.error("Transcription error %s: %s", resp.status, data)
            return None
        return (data.get("text") or "").strip() or None


# =========================
# TTS
# =========================

async def tts_to_mp3(text: str, voice: str = VOICE_NAME) -> Optional[str]:
    text = normalize_text(text)
    if not text:
        return None

    fd, path = tempfile.mkstemp(suffix=".mp3", dir=DATA_DIR)
    os.close(fd)

    try:
        communicate = edge_tts.Communicate(text, voice)
        await communicate.save(path)
        return path
    except Exception as e:
        logger.exception("TTS error: %s", e)
        try:
            os.remove(path)
        except Exception:
            pass
        return None


async def send_text_and_optional_voice(message: Message, text: str, settings: Dict[str, Any]) -> None:
    for chunk in split_text(text):
        await message.answer(chunk)

    if int(settings.get("voice_mode", 0)) == 1:
        tts_text = normalize_text(text)
        if len(tts_text) > 1200:
            tts_text = tts_text[:1200] + "..."
        audio_path = await tts_to_mp3(tts_text)
        if audio_path:
            try:
                await message.answer_audio(FSInputFile(audio_path), caption="Голосовой ответ")
            finally:
                try:
                    os.remove(audio_path)
                except Exception:
                    pass


async def send_long_text(chat_id: int, text: str) -> None:
    for chunk in split_text(text):
        await bot.send_message(chat_id, chunk)


# =========================
# ЛЁГКИЙ ПОИСК ПО ПАМЯТИ (без ML-зависимостей)
# =========================
#
# Раньше здесь использовалась sentence-transformers (+ torch) для смыслового
# поиска. На Railway это даёт огромный вес сборки и долгий деплой, поэтому
# заменено на keyword-скоринг с бонусом за совпадение биграмм (пар слов подряд)
# — это даёт неплохое приближение к "смысловому" совпадению фраз без единой
# ML-зависимости и без скачивания моделей при старте.

def semantic_search_sync(query: str, items: List[Dict[str, Any]], top_k: int = 5) -> List[Dict[str, Any]]:
    query_norm = query.lower().strip()
    tokens = [t for t in query_norm.split() if len(t) > 2]

    if not tokens or not items:
        return []

    scored = []
    for item in items:
        text_lower = item["text"].lower()
        score = 0.0

        for token in tokens:
            if token in text_lower:
                score += 1.0

        for i in range(len(tokens) - 1):
            bigram = f"{tokens[i]} {tokens[i + 1]}"
            if bigram in text_lower:
                score += 1.5

        if score > 0:
            normalized_score = score / (1 + len(text_lower) / 500)
            scored.append({**item, "score": round(normalized_score, 3)})

    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored[:top_k]


async def search_knowledge(query: str, user_id: int, top_k: int = 5) -> List[Dict[str, Any]]:
    memories = await get_memories(user_id, limit=100)
    notes = await get_notes(user_id, limit=100)
    tasks = await get_tasks(user_id, status="open", limit=100)

    items: List[Dict[str, Any]] = []

    for row in memories:
        items.append({
            "source": "memory",
            "id": row["id"],
            "text": row["content"],
        })

    for row in notes:
        items.append({
            "source": "note",
            "id": row["id"],
            "text": f"{row['title']}: {row['content']}",
        })

    for row in tasks:
        due = f" | due: {row['due_at']}" if row.get("due_at") else ""
        items.append({
            "source": "task",
            "id": row["id"],
            "text": f"{row['text']} | status: {row['status']}{due}",
        })

    if not items:
        return []

    return await asyncio.to_thread(semantic_search_sync, query, items, top_k)


# =========================
# WEATHER
# =========================

async def qweather_lookup(city: str) -> Optional[Dict[str, Any]]:
    if not QWEATHER_KEY:
        return None

    session = await get_http_session()
    url = f"https://geoapi.qweather.com/v2/city/lookup?location={quote_plus(city)}&key={QWEATHER_KEY}"

    async with session.get(url) as resp:
        data = await resp.json(content_type=None)
        if resp.status >= 400:
            logger.error("QWeather lookup error %s: %s", resp.status, data)
            return None

        locs = data.get("location") or []
        if not locs:
            return None
        return locs[0]


async def qweather_now(location_id: str) -> Optional[Dict[str, Any]]:
    if not QWEATHER_KEY:
        return None

    session = await get_http_session()
    url = f"https://devapi.qweather.com/v7/weather/now?location={location_id}&key={QWEATHER_KEY}"

    async with session.get(url) as resp:
        data = await resp.json(content_type=None)
        if resp.status >= 400:
            logger.error("QWeather now error %s: %s", resp.status, data)
            return None

        now = data.get("now")
        if not now:
            return None
        return now


async def get_weather_text(city: str) -> str:
    loc = await qweather_lookup(city)
    if not loc:
        return f"Не нашёл город: {city}"

    now = await qweather_now(loc["id"])
    if not now:
        return f"Не удалось получить погоду для {loc['name']}."

    return (
        f"Погода в {loc['name']}:\n"
        f"• Сейчас: {now.get('text', '—')}\n"
        f"• Температура: {now.get('temp', '—')}°C\n"
        f"• Ощущается как: {now.get('feelsLike', '—')}°C\n"
        f"• Влажность: {now.get('humidity', '—')}%\n"
        f"• Ветер: {now.get('windDir', '—')} {now.get('windScale', '—')} м/с\n"
        f"• Видимость: {now.get('vis', '—')} км"
    )


# =========================
# DOCUMENTS
# =========================

def extract_text_from_pdf(data: bytes) -> str:
    reader = PdfReader(BytesIO(data))
    parts = []
    for page in reader.pages:
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            continue
    return normalize_text("\n".join(parts))


def extract_text_from_docx(data: bytes) -> str:
    doc = DocxDocument(BytesIO(data))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return normalize_text("\n".join(parts))


def extract_text_from_xlsx(data: bytes) -> str:
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    parts = []
    for ws in wb.worksheets:
        parts.append(f"[Sheet: {ws.title}]")
        row_count = 0
        for row in ws.iter_rows(values_only=True):
            row_count += 1
            if row_count > 200:
                parts.append("... (sheet truncated)")
                break
            cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cells:
                parts.append(" | ".join(cells))
    return normalize_text("\n".join(parts))


def extract_text_from_txt(data: bytes) -> str:
    for encoding in ("utf-8", "cp1251", "latin-1"):
        try:
            return normalize_text(data.decode(encoding, errors="ignore"))
        except Exception:
            continue
    return normalize_text(data.decode("utf-8", errors="ignore"))


async def summarize_text(text: str, title: str = "") -> str:
    text = clip(text, 14000)
    system = (
        "Ты кратко и точно суммируешь документы по-русски. "
        "Дай 5-10 буллетов, потом 1 короткий вывод. "
        "Не выдумывай факты."
    )
    user = f"Название: {title}\n\nТекст:\n{text}"
    return await groq_chat(
        [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        model=GROQ_MODEL_FAST,
        temperature=0.2,
    )


async def download_telegram_file(file_id: str) -> bytes:
    tg_file = await bot.get_file(file_id)
    buf = BytesIO()
    await bot.download_file(tg_file.file_path, destination=buf)
    return buf.getvalue()


def detect_file_type(filename: str) -> str:
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    return ext


# =========================
# IMAGE ANALYSIS
# =========================

async def analyze_image_with_groq(image_bytes: bytes, caption: str = "") -> str:
    if not GROQ_API_KEY:
        return "GROQ_API_KEY не задан."

    session = await get_http_session()
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    data_url = f"data:image/jpeg;base64,{b64}"

    payload = {
        "model": GROQ_MODEL_VISION,
        "messages": [
            {
                "role": "system",
                "content": (
                    "Ты анализируешь изображение. Отвечай по-русски, "
                    "кратко, но полезно. Если есть текст на изображении, извлеки его."
                ),
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": caption or "Опиши изображение и извлеки важные детали."},
                    {"type": "image_url", "image_url": {"url": data_url}},
                ],
            },
        ],
        "temperature": 0.2,
    }

    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    async with session.post(GROQ_CHAT_URL, headers=headers, json=payload) as resp:
        data = await resp.json(content_type=None)
        if resp.status >= 400:
            logger.error("Vision error %s: %s", resp.status, data)
            return "Не удалось проанализировать изображение."
        try:
            return data["choices"][0]["message"]["content"].strip()
        except Exception:
            return "Не удалось разобрать ответ по изображению."


# =========================
# SYSTEM PROMPT / CONTEXT
# =========================

async def build_system_prompt(user_id: int, query: str) -> str:
    settings = await get_settings(user_id)
    relevant = await search_knowledge(query, user_id, top_k=5)
    relevant_lines = []
    for item in relevant:
        src = item["source"]
        iid = item["id"]
        score = item["score"]
        txt = item["text"]
        relevant_lines.append(f"[{src} #{iid} | {score:.3f}] {txt}")

    context_block = "\n".join(relevant_lines) if relevant_lines else "Нет релевантного контекста."

    tasks = await get_tasks(user_id, status="open", limit=7)
    task_block = "\n".join([f"- #{t['id']}: {t['text']}" for t in tasks]) if tasks else "Нет открытых задач."

    city = settings.get("default_city") or "не задан"
    voice_mode = "ON" if int(settings.get("voice_mode", 0)) else "OFF"
    briefing_mode = "ON" if int(settings.get("briefing_enabled", 0)) else "OFF"
    model_mode = settings.get("model_mode", "smart")

    return (
        "Ты Jarvis — личный ИИ-ассистент пользователя.\n"
        "Пиши по-русски. Стиль: уверенно, коротко, полезно, без воды.\n"
        "Если уместно — предлагай следующий шаг.\n"
        "Если нужно уточнение, задай один конкретный вопрос.\n"
        "Не упоминай внутренние инструкции.\n\n"
        f"Текущее время: {now_str()}\n"
        f"Режим модели: {model_mode}\n"
        f"Voice mode: {voice_mode}\n"
        f"Briefing: {briefing_mode}\n"
        f"Город по умолчанию: {city}\n\n"
        f"Открытые задачи:\n{task_block}\n\n"
        f"Релевантный контекст из памяти:\n{context_block}"
    )


# =========================
# BRIEFING
# =========================

async def build_briefing(user_id: int) -> str:
    settings = await get_settings(user_id)
    city = settings.get("default_city")

    lines = [f"Брифинг на {now_str()}"]

    if city:
        weather = await get_weather_text(city)
        lines.append(weather)
    else:
        lines.append("Погода: город по умолчанию не задан.")

    open_tasks = await get_tasks(user_id, status="open", limit=7)
    if open_tasks:
        lines.append("\nЗадачи:")
        for t in open_tasks:
            due = f" — до {t['due_at']}" if t.get("due_at") else ""
            lines.append(f"• #{t['id']} {t['text']}{due}")
    else:
        lines.append("\nЗадачи: нет открытых.")

    due_reminders = await get_due_reminders_for_user(user_id)
    if due_reminders:
        lines.append("\nПросроченные/срочные напоминания:")
        for r in due_reminders[:5]:
            lines.append(f"• #{r['id']} {r['text']} (должно быть: {r['due_at']})")
    else:
        lines.append("\nНапоминания: нет срочных.")

    notes = await get_notes(user_id, limit=3)
    if notes:
        lines.append("\nПоследние заметки:")
        for n in notes[:3]:
            preview = n["content"][:150].replace("\n", " ")
            lines.append(f"• #{n['id']} {n['title']}: {preview}")

    return "\n".join(lines)


# =========================
# REMINDERS
# =========================

def parse_relative_due(expr: str) -> Optional[datetime]:
    """
    Supported:
    10m, 2h, 1d, 30s
    """
    m = re.match(r"^(\d+)\s*([smhd])$", expr.strip().lower())
    if not m:
        return None
    value = int(m.group(1))
    unit = m.group(2)
    if unit == "s":
        return datetime.now() + timedelta(seconds=value)
    if unit == "m":
        return datetime.now() + timedelta(minutes=value)
    if unit == "h":
        return datetime.now() + timedelta(hours=value)
    if unit == "d":
        return datetime.now() + timedelta(days=value)
    return None


def parse_absolute_due(expr: str) -> Optional[datetime]:
    """
    Supported:
    YYYY-MM-DD HH:MM
    """
    expr = expr.strip()
    for fmt in ("%Y-%m-%d %H:%M", "%d.%m.%Y %H:%M"):
        try:
            return datetime.strptime(expr, fmt)
        except Exception:
            pass
    return None


def parse_reminder_input(raw: str) -> Optional[Dict[str, Any]]:
    raw = raw.strip()
    if not raw:
        return None

    # format: 10m buy milk
    m = re.match(r"^(\d+\s*[smhd])\s+(.+)$", raw, re.I)
    if m:
        due = parse_relative_due(m.group(1))
        if due:
            return {"due": due, "text": m.group(2).strip()}

    # format: 2026-06-25 18:00 buy milk
    m = re.match(r"^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})\s+(.+)$", raw)
    if m:
        due = parse_absolute_due(m.group(1))
        if due:
            return {"due": due, "text": m.group(2).strip()}

    return None


# =========================
# COMMANDS
# =========================

@router.message(Command("start"))
async def cmd_start(message: Message) -> None:
    user_id = message.from_user.id
    await ensure_user(user_id)
    settings = await get_settings(user_id)

    text = (
        "Jarvis online.\n\n"
        "Я умею:\n"
        "• чат с памятью\n"
        "• заметки и поиск\n"
        "• задачи и напоминания\n"
        "• погода\n"
        "• голосовой ввод/ответ\n"
        "• разбор документов\n"
        "• анализ фото\n"
        "• ежедневный брифинг\n\n"
        "Команды:\n"
        "/help — помощь\n"
        "/remember текст — запомнить факт\n"
        "/note текст — сохранить заметку\n"
        "/task текст — добавить задачу\n"
        "/tasks — список задач\n"
        "/done ID — закрыть задачу\n"
        "/remind 10m текст — напоминание\n"
        "/weather город — погода\n"
        "/city город — город по умолчанию\n"
        "/brief — брифинг сейчас\n"
        "/briefing on|off — ежедневный брифинг\n"
        "/voice on|off — голосовые ответы\n"
        "/model fast|smart|vision — режим модели\n"
        "/search запрос — поиск по памяти\n"
        "/profile — настройки\n"
        "/reset — очистить историю\n"
        "/speak текст — озвучить текст\n\n"
        "Можешь просто писать мне как обычному ассистенту."
    )
    await message.answer(text, reply_markup=build_main_keyboard(settings))


@router.message(Command("help"))
async def cmd_help(message: Message) -> None:
    text = (
        "Команды Jarvis:\n\n"
        "/remember текст — сохранить в память\n"
        "/memory — показать память\n"
        "/forget ID — удалить память\n"
        "/note текст — сохранить заметку\n"
        "/notes — список заметок\n"
        "/task текст — добавить задачу\n"
        "/tasks — список задач\n"
        "/done ID — завершить задачу\n"
        "/remind 10m текст — напоминание\n"
        "/weather город — погода\n"
        "/city город — город по умолчанию\n"
        "/brief — брифинг сейчас\n"
        "/briefing on|off — ежедневный брифинг\n"
        "/voice on|off — голосовые ответы\n"
        "/model fast|smart|vision — режим модели\n"
        "/search запрос — поиск по памяти\n"
        "/profile — настройки\n"
        "/reset — очистить историю\n"
        "/speak текст — озвучить текст\n\n"
        "Также поддерживаются:\n"
        "• голосовые сообщения\n"
        "• фото\n"
        "• документы: PDF, DOCX, XLSX, TXT"
    )
    await message.answer(text)


@router.message(Command("profile"))
async def cmd_profile(message: Message) -> None:
    user_id = message.from_user.id
    settings = await get_settings(user_id)
    text = (
        "Профиль:\n"
        f"• voice_mode: {settings['voice_mode']}\n"
        f"• model_mode: {settings['model_mode']}\n"
        f"• briefing_enabled: {settings['briefing_enabled']}\n"
        f"• briefing_time: {settings['briefing_time']}\n"
        f"• default_city: {settings.get('default_city') or 'не задан'}\n"
    )
    await message.answer(text)


@router.message(Command("reset"))
async def cmd_reset(message: Message) -> None:
    user_id = message.from_user.id
    await clear_history(user_id)
    await message.answer("История очищена.")


@router.message(Command("voice"))
async def cmd_voice(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    arg = (command.args or "").strip().lower()
    if arg not in {"on", "off"}:
        await message.answer("Используй: /voice on или /voice off")
        return
    await set_setting(user_id, "voice_mode", 1 if arg == "on" else 0)
    await message.answer(f"Voice mode: {arg.upper()}")


@router.message(Command("model"))
async def cmd_model(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    arg = (command.args or "").strip().lower()
    if arg not in {"fast", "smart", "vision"}:
        await message.answer("Используй: /model fast, /model smart или /model vision")
        return
    await set_setting(user_id, "model_mode", arg)
    await message.answer(f"Model mode: {arg}")


@router.message(Command("briefing"))
async def cmd_briefing(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    arg = (command.args or "").strip().lower()
    if arg not in {"on", "off"}:
        await message.answer("Используй: /briefing on или /briefing off")
        return
    await set_setting(user_id, "briefing_enabled", 1 if arg == "on" else 0)
    if arg == "on":
        await message.answer("Ежедневный брифинг включён. Время по умолчанию: 09:00.")
    else:
        await message.answer("Ежедневный брифинг выключен.")


@router.message(Command("brief"))
async def cmd_brief(message: Message) -> None:
    user_id = message.from_user.id
    briefing = await build_briefing(user_id)
    await message.answer(briefing)


@router.message(Command("city"))
async def cmd_city(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    city = (command.args or "").strip()
    if not city:
        await message.answer("Используй: /city Москва")
        return
    await set_setting(user_id, "default_city", city)
    await message.answer(f"Город по умолчанию установлен: {city}")


@router.message(Command("weather"))
async def cmd_weather(message: Message, command: CommandObject) -> None:
    city = (command.args or "").strip()
    if not city:
        settings = await get_settings(message.from_user.id)
        city = settings.get("default_city") or ""
    if not city:
        await message.answer("Укажи город: /weather Москва")
        return
    text = await get_weather_text(city)
    await message.answer(text)


@router.message(Command("speak"))
async def cmd_speak(message: Message, command: CommandObject) -> None:
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /speak текст")
        return
    audio_path = await tts_to_mp3(text)
    if not audio_path:
        await message.answer("Не смог озвучить текст.")
        return
    try:
        await message.answer_audio(FSInputFile(audio_path), caption="Озвучка")
    finally:
        try:
            os.remove(audio_path)
        except Exception:
            pass


@router.message(Command("remember"))
async def cmd_remember(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /remember текст")
        return
    memory_id = await add_memory(user_id, text)
    await message.answer(f"Запомнил. ID памяти: {memory_id}")


@router.message(Command("memory"))
async def cmd_memory(message: Message) -> None:
    user_id = message.from_user.id
    rows = await get_memories(user_id, limit=20)
    if not rows:
        await message.answer("Память пока пустая.")
        return
    text = "Память:\n" + "\n".join([f"• #{r['id']} {r['content']}" for r in rows])
    await message.answer(text)


@router.message(Command("forget"))
async def cmd_forget(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    arg = (command.args or "").strip()
    if not arg.isdigit():
        await message.answer("Используй: /forget ID")
        return
    await delete_memory(user_id, int(arg))
    await message.answer("Удалил из памяти.")


@router.message(Command("note"))
async def cmd_note(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /note текст")
        return
    title = text[:40].strip()
    note_id = await add_note(user_id, title=title, content=text)
    await message.answer(f"Заметка сохранена. ID: {note_id}")


@router.message(Command("notes"))
async def cmd_notes(message: Message) -> None:
    user_id = message.from_user.id
    rows = await get_notes(user_id, limit=20)
    if not rows:
        await message.answer("Заметок пока нет.")
        return
    text = "Заметки:\n" + "\n".join(
        [f"• #{r['id']} {r['title']} — {r['content'][:120]}" for r in rows]
    )
    await message.answer(text)


@router.message(Command("task"))
async def cmd_task(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /task текст")
        return
    task_id = await add_task(user_id, text)
    await message.answer(f"Задача добавлена. ID: {task_id}")


@router.message(Command("tasks"))
async def cmd_tasks(message: Message) -> None:
    user_id = message.from_user.id
    rows = await get_tasks(user_id, status="open", limit=20)
    if not rows:
        await message.answer("Открытых задач нет.")
        return
    text = "Задачи:\n" + "\n".join(
        [
            f"• #{r['id']} {r['text']}" + (f" — due: {r['due_at']}" if r.get("due_at") else "")
            for r in rows
        ]
    )
    await message.answer(text)


@router.message(Command("done"))
async def cmd_done(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    arg = (command.args or "").strip()
    if not arg.isdigit():
        await message.answer("Используй: /done ID")
        return
    ok = await mark_task_done(user_id, int(arg))
    await message.answer("Задача закрыта." if ok else "Задача не найдена.")


@router.message(Command("remind"))
async def cmd_remind(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /remind 10m текст\nИли: /remind 2026-06-25 18:00 текст")
        return

    parsed = parse_reminder_input(raw)
    if not parsed:
        await message.answer("Не понял формат. Пример: /remind 10m купить воду")
        return

    due = parsed["due"].isoformat(timespec="minutes")
    text = parsed["text"]
    rid = await add_reminder(user_id, message.chat.id, text, due)
    await message.answer(f"Напоминание сохранено. ID: {rid}\nВремя: {due}")


@router.message(Command("search"))
async def cmd_search(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    query = (command.args or "").strip()
    if not query:
        await message.answer("Используй: /search запрос")
        return
    results = await search_knowledge(query, user_id, top_k=5)
    if not results:
        await message.answer("Ничего не нашёл.")
        return

    lines = ["Результаты поиска:"]
    for r in results:
        lines.append(f"• [{r['source']} #{r['id']}] {r['text'][:220]} (score={r['score']:.3f})")
    await message.answer("\n".join(lines))


@router.message(Command("ask"))
async def cmd_ask(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    query = (command.args or "").strip()
    if not query:
        await message.answer("Используй: /ask вопрос")
        return
    await process_user_text(message, query)


# =========================
# CALLBACKS
# =========================

@router.callback_query(F.data == "jarvis_voice")
async def cb_voice(callback: CallbackQuery) -> None:
    user_id = callback.from_user.id
    settings = await get_settings(user_id)
    new_value = 0 if int(settings.get("voice_mode", 0)) else 1
    await set_setting(user_id, "voice_mode", new_value)
    await callback.answer("Готово")
    await callback.message.answer(f"Voice mode: {'ON' if new_value else 'OFF'}")


@router.callback_query(F.data == "jarvis_briefing")
async def cb_briefing(callback: CallbackQuery) -> None:
    user_id = callback.from_user.id
    settings = await get_settings(user_id)
    new_value = 0 if int(settings.get("briefing_enabled", 0)) else 1
    await set_setting(user_id, "briefing_enabled", new_value)
    await callback.answer("Готово")
    await callback.message.answer(f"Briefing: {'ON' if new_value else 'OFF'}")


@router.callback_query(F.data == "jarvis_weather")
async def cb_weather(callback: CallbackQuery) -> None:
    await callback.answer()
    settings = await get_settings(callback.from_user.id)
    city = settings.get("default_city")
    if city:
        text = await get_weather_text(city)
        await callback.message.answer(text)
    else:
        await callback.message.answer("Сначала задай город: /city Москва")


@router.callback_query(F.data == "jarvis_tasks")
async def cb_tasks(callback: CallbackQuery) -> None:
    await callback.answer()
    rows = await get_tasks(callback.from_user.id, status="open", limit=10)
    if not rows:
        await callback.message.answer("Открытых задач нет.")
        return
    text = "Задачи:\n" + "\n".join([f"• #{r['id']} {r['text']}" for r in rows])
    await callback.message.answer(text)


@router.callback_query(F.data == "jarvis_memory")
async def cb_memory(callback: CallbackQuery) -> None:
    await callback.answer()
    rows = await get_memories(callback.from_user.id, limit=10)
    if not rows:
        await callback.message.answer("Память пока пустая.")
        return
    text = "Память:\n" + "\n".join([f"• #{r['id']} {r['content']}" for r in rows])
    await callback.message.answer(text)


@router.callback_query(F.data == "jarvis_remind")
async def cb_remind(callback: CallbackQuery) -> None:
    await callback.answer()
    await callback.message.answer("Создай напоминание командой:\n/remind 10m купить воду")


# =========================
# TEXT / CHAT
# =========================

async def generate_jarvis_reply(user_id: int, text: str) -> Optional[str]:
    """
    Ядро логики Jarvis: принимает user_id и текст, возвращает текст ответа
    (или None, если это был "быстрый интент" без ответа, например /remember).
    Не зависит от Telegram — используется и ботом, и HTTP /jarvis эндпоинтом.
    """
    await ensure_user(user_id)
    text = normalize_text(text)
    if not text:
        return None

    # quick intents
    if text.lower().startswith("запомни "):
        content = text.split(" ", 1)[1].strip()
        if content:
            mid = await add_memory(user_id, content)
            return f"Запомнил. ID памяти: {mid}"
        return None

    if text.lower().startswith("заметка "):
        content = text.split(" ", 1)[1].strip()
        if content:
            nid = await add_note(user_id, title=content[:40], content=content)
            return f"Заметка сохранена. ID: {nid}"
        return None

    settings = await get_settings(user_id)

    # save user message
    await save_message(user_id, "user", text)

    # build context
    system_prompt = await build_system_prompt(user_id, text)
    history = await get_recent_messages(user_id, limit=12)

    model = select_model(settings)

    messages: List[Dict[str, Any]] = [{"role": "system", "content": system_prompt}]
    messages.extend(history)

    answer = await groq_chat(messages, model=model, temperature=0.5)
    if not answer:
        answer = "Пока не смог сформировать ответ."

    await save_message(user_id, "assistant", answer)
    return answer


async def process_user_text(message: Message, text: str) -> None:
    user_id = message.from_user.id
    settings = await get_settings(user_id)

    answer = await generate_jarvis_reply(user_id, text)
    if answer is None:
        return

    await send_text_and_optional_voice(message, answer, settings)


@router.message(F.text)
async def on_text(message: Message) -> None:
    text = (message.text or "").strip()
    if not text:
        return
    if text.startswith("/"):
        return
    await process_user_text(message, text)


@router.message(F.voice)
async def on_voice(message: Message) -> None:
    try:
        voice = message.voice
        tg_file = await bot.get_file(voice.file_id)
        buf = BytesIO()
        await bot.download_file(tg_file.file_path, destination=buf)
        transcript = await groq_transcribe(buf.getvalue(), filename="voice.ogg")
        if not transcript:
            await message.answer("Не смог распознать голосовое сообщение.")
            return

        await message.answer(f"Распознал:\n{transcript}")
        await process_user_text(message, transcript)
    except Exception as e:
        logger.exception("Voice handling error: %s", e)
        await message.answer("Ошибка при обработке голосового сообщения.")


@router.message(F.photo)
async def on_photo(message: Message) -> None:
    try:
        photo = message.photo[-1]
        tg_file = await bot.get_file(photo.file_id)
        buf = BytesIO()
        await bot.download_file(tg_file.file_path, destination=buf)
        caption = message.caption or "Опиши изображение и извлеки важные детали."
        result = await analyze_image_with_groq(buf.getvalue(), caption=caption)
        await message.answer(result)
    except Exception as e:
        logger.exception("Photo handling error: %s", e)
        await message.answer("Ошибка при анализе изображения.")


@router.message(F.document)
async def on_document(message: Message) -> None:
    try:
        doc = message.document
        filename = doc.file_name or "file"
        ext = detect_file_type(filename)

        tg_file = await bot.get_file(doc.file_id)
        buf = BytesIO()
        await bot.download_file(tg_file.file_path, destination=buf)
        data = buf.getvalue()

        text = ""
        if ext == "pdf":
            text = extract_text_from_pdf(data)
        elif ext == "docx":
            text = extract_text_from_docx(data)
        elif ext in {"xlsx", "xlsm", "xltx", "xltm"}:
            text = extract_text_from_xlsx(data)
        elif ext in {"txt", "md", "csv", "log"}:
            text = extract_text_from_txt(data)
        else:
            await message.answer(
                "Поддерживаются: PDF, DOCX, XLSX, TXT, MD, CSV, LOG.\n"
                "Если хочешь, я могу позже добавить OCR для изображений и сканов."
            )
            return

        if not text:
            await message.answer("Не смог извлечь текст из файла.")
            return

        summary = await summarize_text(text, title=filename)
        note_title = f"File: {filename}"
        await add_note(message.from_user.id, note_title, summary)
        await message.answer(
            f"Файл обработан: {filename}\n\n"
            f"Краткая сводка:\n{summary}"
        )
    except Exception as e:
        logger.exception("Document handling error: %s", e)
        await message.answer("Ошибка при обработке документа.")


# =========================
# SCHEDULER JOBS
# =========================

async def reminders_job() -> None:
    reminders = await get_due_reminders(limit=100)
    if not reminders:
        return

    for r in reminders:
        try:
            await bot.send_message(
                r["chat_id"],
                f"⏰ Напоминание #{r['id']}:\n{r['text']}\n\nВремя: {r['due_at']}",
            )
            await mark_reminder_sent(r["id"])
        except Exception as e:
            logger.warning("Failed to send reminder #%s: %s", r["id"], e)


async def briefing_job() -> None:
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M")

    users = await db_query(
        """
        SELECT user_id, briefing_time, last_brief_date
        FROM settings
        WHERE briefing_enabled = 1
        """
    )

    for u in users:
        try:
            if u["briefing_time"] != current_time:
                continue
            if u.get("last_brief_date") == today:
                continue

            text = await build_briefing(u["user_id"])
            await bot.send_message(u["user_id"], text)
            await set_setting(u["user_id"], "last_brief_date", today)
        except Exception as e:
            logger.warning("Failed briefing for %s: %s", u["user_id"], e)


# =========================
# HTTP API (для MacroDroid / любого внешнего клиента)
# =========================
#
# Один простой endpoint: POST /jarvis с JSON {"text": "..."}.
# MacroDroid сам распознаёт речь в текст (Android STT) и сам озвучивает
# полученный ответ (Android TTS) — поэтому сюда гоняется только текст,
# никакого аудио по сети передавать не нужно. Это резко проще и быстрее.
#
# Защита простым статическим ключом через заголовок X-API-Key.
# Если JARVIS_API_KEY не задан в переменных окружения — endpoint выключен
# (чтобы случайно не оставить сервер открытым всему интернету).

async def handle_jarvis_request(request: web.Request) -> web.Response:
    if not JARVIS_API_KEY:
        return web.json_response(
            {"error": "JARVIS_API_KEY не задан на сервере — HTTP API выключен."},
            status=503,
        )

    api_key = request.headers.get("X-API-Key", "")
    if api_key != JARVIS_API_KEY:
        return web.json_response({"error": "unauthorized"}, status=401)

    try:
        data = await request.json()
    except Exception:
        return web.json_response({"error": "invalid JSON body"}, status=400)

    text = (data.get("text") or "").strip()
    if not text:
        return web.json_response({"error": "поле 'text' пустое или отсутствует"}, status=400)

    user_id = JARVIS_HTTP_USER_ID
    if not user_id:
        return web.json_response(
            {"error": "JARVIS_HTTP_USER_ID не задан на сервере."},
            status=503,
        )

    try:
        answer = await generate_jarvis_reply(user_id, text)
    except Exception as e:
        logger.exception("HTTP /jarvis error: %s", e)
        return web.json_response({"error": "internal error"}, status=500)

    if answer is None:
        answer = "Готово."

    return web.json_response({"reply": answer})


async def handle_health(request: web.Request) -> web.Response:
    return web.json_response({"status": "ok"})


def build_http_app() -> web.Application:
    app = web.Application()
    app.router.add_post("/jarvis", handle_jarvis_request)
    app.router.add_get("/health", handle_health)
    return app


async def start_http_server() -> web.AppRunner:
    port = int(os.environ.get("PORT", "8080"))
    app = build_http_app()
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, host="0.0.0.0", port=port)
    await site.start()
    logger.info("HTTP API запущен на порту %s (/jarvis, /health).", port)
    return runner


# =========================
# STARTUP / SHUTDOWN
# =========================

async def on_startup() -> None:
    init_db()
    scheduler.add_job(reminders_job, "interval", seconds=30, id="reminders_job", replace_existing=True)
    scheduler.add_job(briefing_job, "interval", seconds=60, id="briefing_job", replace_existing=True)
    scheduler.start()
    logger.info("Jarvis started.")


async def on_shutdown() -> None:
    try:
        scheduler.shutdown(wait=False)
    except Exception:
        pass

    global HTTP_SESSION
    if HTTP_SESSION and not HTTP_SESSION.closed:
        await HTTP_SESSION.close()

    await bot.session.close()
    logger.info("Jarvis stopped.")


async def main() -> None:
    await on_startup()
    http_runner = await start_http_server()
    try:
        await dp.start_polling(bot)
    finally:
        await http_runner.cleanup()
        await on_shutdown()


if __name__ == "__main__":
    asyncio.run(main())
