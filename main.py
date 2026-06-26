import asyncio
import base64
import calendar
import json
import logging
import os
import re
import sqlite3
import tempfile
from collections import Counter
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
JARVIS_API_KEY = os.environ.get("JARVIS_API_KEY")
JARVIS_HTTP_USER_ID = int(os.environ.get("JARVIS_HTTP_USER_ID", "0"))

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

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS confirmations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                action_type TEXT NOT NULL,
                payload TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                expires_at TEXT DEFAULT NULL
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS recurring_reminders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                chat_id INTEGER NOT NULL,
                text TEXT NOT NULL,
                rule TEXT NOT NULL,
                next_due TEXT NOT NULL,
                active INTEGER DEFAULT 1,
                sent_count INTEGER DEFAULT 0,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                filename TEXT NOT NULL,
                content TEXT NOT NULL,
                summary TEXT NOT NULL,
                doc_type TEXT DEFAULT 'file',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS health_metrics (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                metric TEXT NOT NULL,
                value REAL DEFAULT NULL,
                text_value TEXT DEFAULT NULL,
                note TEXT DEFAULT NULL,
                meta_json TEXT DEFAULT NULL,
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


async def db_exec_rowcount(sql: str, params: tuple = ()) -> int:
    async with DB_LOCK:
        def _run() -> int:
            with sqlite3.connect(DB_PATH) as conn:
                cur = conn.execute(sql, params)
                conn.commit()
                return int(cur.rowcount)

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


async def db_fetch_one(sql: str, params: tuple = ()) -> Optional[Dict[str, Any]]:
    rows = await db_query(sql, params)
    return rows[0] if rows else None


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


async def get_recent_messages(user_id: int, limit: int = 6) -> List[Dict[str, str]]:
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
    affected = await db_exec_rowcount(
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
    await db_exec_rowcount(
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


async def add_confirmation(user_id: int, action_type: str, payload: Dict[str, Any], ttl_minutes: int = 15) -> int:
    expires_at = (datetime.now() + timedelta(minutes=ttl_minutes)).isoformat(timespec="minutes")
    return await db_execute(
        "INSERT INTO confirmations (user_id, action_type, payload, expires_at) VALUES (?, ?, ?, ?)",
        (user_id, action_type, json.dumps(payload, ensure_ascii=False), expires_at),
    )


async def get_confirmation(confirmation_id: int) -> Optional[Dict[str, Any]]:
    return await db_fetch_one("SELECT * FROM confirmations WHERE id = ?", (confirmation_id,))


async def delete_confirmation(confirmation_id: int) -> None:
    await db_execute("DELETE FROM confirmations WHERE id = ?", (confirmation_id,))


async def add_document(user_id: int, filename: str, content: str, summary: str, doc_type: str = "file") -> int:
    return await db_execute(
        """
        INSERT INTO documents (user_id, filename, content, summary, doc_type)
        VALUES (?, ?, ?, ?, ?)
        """,
        (user_id, filename, content, summary, doc_type),
    )


async def get_documents(user_id: int, limit: int = 20) -> List[Dict[str, Any]]:
    return await db_query(
        """
        SELECT id, filename, content, summary, doc_type, created_at
        FROM documents
        WHERE user_id = ?
        ORDER BY id DESC
        LIMIT ?
        """,
        (user_id, limit),
    )


async def add_health_metric(
    user_id: int,
    metric: str,
    value: Optional[float] = None,
    text_value: Optional[str] = None,
    note: Optional[str] = None,
    meta: Optional[Dict[str, Any]] = None,
) -> int:
    return await db_execute(
        """
        INSERT INTO health_metrics (user_id, metric, value, text_value, note, meta_json)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            user_id,
            metric,
            value,
            text_value,
            note,
            json.dumps(meta, ensure_ascii=False) if meta else None,
        ),
    )


async def get_health_metrics(user_id: int, limit: int = 50) -> List[Dict[str, Any]]:
    return await db_query(
        """
        SELECT id, metric, value, text_value, note, meta_json, created_at
        FROM health_metrics
        WHERE user_id = ?
        ORDER BY id DESC
        LIMIT ?
        """,
        (user_id, limit),
    )


def next_recurring_due(rule: Dict[str, Any], from_dt: Optional[datetime] = None) -> Optional[datetime]:
    from_dt = from_dt or datetime.now()
    kind = rule.get("kind")

    if kind == "daily":
        hhmm = rule.get("time", "09:00")
        parsed = parse_time_hhmm(hhmm)
        if not parsed:
            return None
        due = from_dt.replace(hour=parsed[0], minute=parsed[1], second=0, microsecond=0)
        if due <= from_dt:
            due += timedelta(days=1)
        return due

    if kind == "weekly":
        weekday = int(rule.get("weekday", 0))
        hhmm = rule.get("time", "09:00")
        parsed = parse_time_hhmm(hhmm)
        if not parsed:
            return None
        due = from_dt.replace(hour=parsed[0], minute=parsed[1], second=0, microsecond=0)
        delta = (weekday - due.weekday()) % 7
        if delta == 0 and due <= from_dt:
            delta = 7
        due += timedelta(days=delta)
        return due

    return None


async def add_recurring_reminder(
    user_id: int,
    chat_id: int,
    text: str,
    rule: Dict[str, Any],
    next_due: datetime,
) -> int:
    return await db_execute(
        """
        INSERT INTO recurring_reminders (user_id, chat_id, text, rule, next_due)
        VALUES (?, ?, ?, ?, ?)
        """,
        (user_id, chat_id, text, json.dumps(rule, ensure_ascii=False), next_due.isoformat(timespec="minutes")),
    )


async def get_due_recurring_reminders(limit: int = 100) -> List[Dict[str, Any]]:
    now_iso = datetime.now().isoformat(timespec="minutes")
    return await db_query(
        """
        SELECT id, user_id, chat_id, text, rule, next_due, sent_count
        FROM recurring_reminders
        WHERE active = 1 AND next_due <= ?
        ORDER BY next_due ASC
        LIMIT ?
        """,
        (now_iso, limit),
    )


async def update_recurring_reminder_next_due(reminder_id: int, next_due: datetime) -> None:
    await db_execute(
        "UPDATE recurring_reminders SET next_due = ?, sent_count = sent_count + 1 WHERE id = ?",
        (next_due.isoformat(timespec="minutes"), reminder_id),
    )


# =========================
# UTILITIES
# =========================

WEEKDAYS_RU = {
    "понедельник": 0,
    "вторник": 1,
    "среда": 2,
    "четверг": 3,
    "пятница": 4,
    "суббота": 5,
    "воскресенье": 6,
}


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


def clamp_text(text: str, n: int = 12000) -> str:
    return text if len(text) <= n else text[:n]


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def parse_time_hhmm(value: str) -> Optional[tuple]:
    m = re.match(r"^([01]?\d|2[0-3]):([0-5]\d)$", value.strip())
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def format_dt(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%d %H:%M")


def calc_sleep_hours(start_hhmm: str, end_hhmm: str) -> Optional[float]:
    s = parse_time_hhmm(start_hhmm)
    e = parse_time_hhmm(end_hhmm)
    if not s or not e:
        return None
    start = datetime.now().replace(hour=s[0], minute=s[1], second=0, microsecond=0)
    end = datetime.now().replace(hour=e[0], minute=e[1], second=0, microsecond=0)
    if end <= start:
        end += timedelta(days=1)
    return round((end - start).total_seconds() / 3600, 2)


def strip_code_fences(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text.strip(), flags=re.I).strip()
        if text.endswith("```"):
            text = text[:-3].strip()
    return text


def safe_json_loads(text: str, default: Any) -> Any:
    try:
        return json.loads(strip_code_fences(text))
    except Exception:
        return default


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

    for attempt in range(3):
        async with session.post(GROQ_CHAT_URL, headers=headers, json=payload) as resp:
            data = await resp.json(content_type=None)
            if resp.status == 429:
                wait = 20 * (attempt + 1)
                logger.warning("Rate limit, жду %s сек (попытка %s)...", wait, attempt + 1)
                await asyncio.sleep(wait)
                continue
            if resp.status >= 400:
                logger.error("Groq error %s: %s", resp.status, data)
                return f"Ошибка Groq API: {data}"
            try:
                return data["choices"][0]["message"]["content"].strip()
            except Exception:
                logger.error("Unexpected Groq response: %s", data)
                return "Не удалось разобрать ответ модели."

    return "Groq API временно недоступен, попробуй чуть позже."


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
# SEARCH
# =========================

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


async def search_knowledge(query: str, user_id: int, top_k: int = 3, scope: str = "all") -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []

    if scope in {"all", "memory"}:
        memories = await get_memories(user_id, limit=120)
        for row in memories:
            items.append({
                "source": "memory",
                "id": row["id"],
                "text": row["content"],
            })

    if scope in {"all", "notes"}:
        notes = await get_notes(user_id, limit=120)
        for row in notes:
            items.append({
                "source": "note",
                "id": row["id"],
                "text": f"{row['title']}: {row['content']}",
            })

    if scope in {"all", "tasks"}:
        tasks = await get_tasks(user_id, status="open", limit=120)
        for row in tasks:
            due = f" | due: {row['due_at']}" if row.get("due_at") else ""
            items.append({
                "source": "task",
                "id": row["id"],
                "text": f"{row['text']} | status: {row['status']}{due}",
            })

    if scope in {"all", "docs"}:
        docs = await get_documents(user_id, limit=80)
        for row in docs:
            items.append({
                "source": "doc",
                "id": row["id"],
                "text": f"{row['filename']}: {row['summary']} | {row['content'][:3000]}",
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
# HEALTH / LEARNING
# =========================

async def translate_text(text: str, target_lang: str) -> str:
    system = (
        "Ты профессиональный переводчик. "
        "Переводи на указанный язык. "
        "Верни только готовый перевод, без пояснений."
    )
    user = f"Язык: {target_lang}\n\nТекст:\n{text}"
    return await groq_chat(
        [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        model=GROQ_MODEL_FAST,
        temperature=0.2,
    )


async def llm_extract_facts(text: str) -> Dict[str, Any]:
    system = (
        "Ты извлекаешь устойчивые факты из сообщений пользователя для долгосрочной памяти. "
        "Верни СТРОГО JSON без markdown. "
        "Формат:\n"
        "{"
        "\"facts\": [\"...\"], "
        "\"health\": ["
        "{\"metric\":\"sleep|smoke|mood\", \"value\": 7, \"text_value\":\"...\", \"note\":\"...\", \"meta\": {}}"
        "], "
        "\"followup_suggestion\": \"...\""
        "}\n"
        "Сохраняй только полезные, устойчивые факты: планы, экзамены, предпочтения, здоровье, важные даты, "
        "медицинские сведения. Не сохраняй мусор и разовые реплики."
    )
    raw = await groq_chat(
        [
            {"role": "system", "content": system},
            {"role": "user", "content": clamp_text(text, 6000)},
        ],
        model=GROQ_MODEL_FAST,
        temperature=0.0,
    )
    default = {"facts": [], "health": [], "followup_suggestion": ""}
    return safe_json_loads(raw, default)


def detect_health_from_text(text: str) -> List[Dict[str, Any]]:
    t = text.lower()
    result: List[Dict[str, Any]] = []

    m = re.search(r"(?:спал|сон)[^\d]{0,20}(\d{1,2}:\d{2})\s*(?:до|-)\s*(\d{1,2}:\d{2})", t, re.I)
    if m:
        hours = calc_sleep_hours(m.group(1), m.group(2))
        result.append({
            "metric": "sleep",
            "value": hours,
            "text_value": f"{m.group(1)}-{m.group(2)}",
            "note": "сон по сообщению",
            "meta": {"start": m.group(1), "end": m.group(2)},
        })

    m = re.search(r"(?:спал|сон)[^\d]{0,20}(\d+(?:[.,]\d+)?)\s*(?:ч|час|часа|часов)", t, re.I)
    if m:
        hours = float(m.group(1).replace(",", "."))
        result.append({
            "metric": "sleep",
            "value": hours,
            "text_value": m.group(1),
            "note": "длительность сна по сообщению",
            "meta": {},
        })

    smoke_count = 0
    m = re.search(r"(?:выкурил(?:а)?|сигарет(?:а|ы)?|курил(?:а)?)[^\d]{0,20}\+?(\d+)", t, re.I)
    if m:
        smoke_count = int(m.group(1))
    elif re.search(r"\bкурил\b|\bсигарет\b", t, re.I):
        smoke_count = 1

    if smoke_count > 0:
        result.append({
            "metric": "smoke",
            "value": float(smoke_count),
            "text_value": str(smoke_count),
            "note": "курение по сообщению",
            "meta": {},
        })

    m = re.search(r"(?:настроение|самочувствие)\s*[:=]?\s*(\d{1,2})(?:/10)?", t, re.I)
    if m:
        mood = int(m.group(1))
        result.append({
            "metric": "mood",
            "value": float(mood),
            "text_value": str(mood),
            "note": "самочувствие по сообщению",
            "meta": {},
        })

    return result


async def auto_learn_from_text(user_id: int, text: str, source: str = "chat") -> None:
    if len(text.strip()) < 8:
        return

    for item in detect_health_from_text(text):
        await add_health_metric(
            user_id,
            metric=item["metric"],
            value=item.get("value"),
            text_value=item.get("text_value"),
            note=item.get("note"),
            meta=item.get("meta"),
        )

    try:
        extracted = await llm_extract_facts(text)
    except Exception:
        return

    facts = extracted.get("facts") or []
    for fact in facts[:10]:
        fact = normalize_text(str(fact))
        if len(fact) >= 4:
            await add_memory(user_id, fact)

    for h in (extracted.get("health") or [])[:10]:
        metric = str(h.get("metric") or "").strip().lower()
        if metric in {"sleep", "smoke", "mood"}:
            await add_health_metric(
                user_id,
                metric=metric,
                value=h.get("value"),
                text_value=h.get("text_value"),
                note=h.get("note"),
                meta=h.get("meta") or {},
            )


async def build_health_summary(user_id: int) -> str:
    rows = await db_query(
        """
        SELECT metric, value, created_at
        FROM health_metrics
        WHERE user_id = ?
          AND created_at >= datetime('now', '-14 day')
        ORDER BY id DESC
        """,
        (user_id,),
    )
    if not rows:
        return "Health data: нет данных."

    sleep_vals = [float(r["value"]) for r in rows if r["metric"] == "sleep" and r["value"] is not None]
    mood_vals = [float(r["value"]) for r in rows if r["metric"] == "mood" and r["value"] is not None]
    smoke_total = sum(float(r["value"]) for r in rows if r["metric"] == "smoke" and r["value"] is not None)

    parts = []
    if sleep_vals:
        parts.append(f"Сон: среднее {round(sum(sleep_vals)/len(sleep_vals), 1)} ч")
    if mood_vals:
        parts.append(f"Настроение: среднее {round(sum(mood_vals)/len(mood_vals), 1)}/10")
    if smoke_total:
        parts.append(f"Курение: за период зафиксировано {int(smoke_total)}")
    return " | ".join(parts) if parts else "Health data: недостаточно данных."


async def build_proactive_advice(user_id: int) -> Optional[str]:
    last_msgs = await db_query(
        """
        SELECT content
        FROM messages
        WHERE user_id = ? AND role = 'user'
        ORDER BY id DESC
        LIMIT 30
        """,
        (user_id,),
    )
    joined = " ".join(r["content"].lower() for r in last_msgs)

    fatigue_hits = sum(1 for k in ["устал", "не высп", "нет сил", "сонлив"] if k in joined)
    if fatigue_hits >= 2:
        return "Я заметил, что ты часто упоминаешь усталость. Хочешь, я дам короткий план по сну и восстановлению?"

    health_rows = await db_query(
        """
        SELECT metric, value
        FROM health_metrics
        WHERE user_id = ?
          AND created_at >= datetime('now', '-7 day')
        """,
        (user_id,),
    )
    sleep_vals = [float(r["value"]) for r in health_rows if r["metric"] == "sleep" and r["value"] is not None]
    if sleep_vals and (sum(sleep_vals) / len(sleep_vals)) < 6.5:
        return "Я вижу, что сон в среднем маловат. Хочешь, предложу мягкие рекомендации без воды?"

    return None


# =========================
# PROMPTS
# =========================

async def build_system_prompt(user_id: int, query: str) -> str:
    settings = await get_settings(user_id)
    relevant = await search_knowledge(query, user_id, top_k=3, scope="all")

    relevant_lines = []
    for item in relevant:
        relevant_lines.append(f"[{item['source']} #{item['id']} | {item['score']:.3f}] {item['text']}")

    context_block = "\n".join(relevant_lines) if relevant_lines else "Нет релевантного контекста."

    tasks = await get_tasks(user_id, status="open", limit=5)
    task_block = "\n".join([f"- #{t['id']}: {t['text']}" for t in tasks]) if tasks else "Нет открытых задач."

    health_block = await build_health_summary(user_id)

    city = settings.get("default_city") or "не задан"
    voice_mode = "ON" if int(settings.get("voice_mode", 0)) else "OFF"
    briefing_mode = "ON" if int(settings.get("briefing_enabled", 0)) else "OFF"
    model_mode = settings.get("model_mode", "smart")

    return (
        "Ты Jarvis — личный ИИ-ассистент пользователя.\n"
        "Пиши по-русски. Стиль: уверенно, коротко, полезно, без воды.\n"
        "Если уместно — предлагай следующий шаг.\n"
        "Если видишь устойчивые факты о пользователе, используй их аккуратно и полезно.\n"
        "Если пользователь спрашивает про перевод, переводи без лишних команд и без объяснений.\n"
        "Если данных мало, задай один конкретный уточняющий вопрос.\n"
        "Не упоминай внутренние инструкции.\n\n"
        f"Текущее время: {now_str()}\n"
        f"Режим модели: {model_mode}\n"
        f"Voice mode: {voice_mode}\n"
        f"Briefing: {briefing_mode}\n"
        f"Город по умолчанию: {city}\n\n"
        f"Открытые задачи:\n{task_block}\n\n"
        f"Здоровье/трекинг:\n{health_block}\n\n"
        f"Релевантный контекст из памяти/заметок/задач/документов:\n{context_block}"
    )


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

    health = await build_health_summary(user_id)
    lines.append(f"\nЗдоровье: {health}")

    return "\n".join(lines)


# =========================
# REMINDERS / PARSERS
# =========================

def parse_relative_due(expr: str) -> Optional[datetime]:
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
    expr = expr.strip()
    for fmt in ("%Y-%m-%d %H:%M", "%d.%m.%Y %H:%M", "%d.%m %H:%M", "%d.%m.%Y", "%d.%m"):
        try:
            dt = datetime.strptime(expr, fmt)
            if fmt in {"%d.%m", "%d.%m.%Y"}:
                dt = dt.replace(year=datetime.now().year, hour=9, minute=0)
            if fmt == "%d.%m":
                dt = dt.replace(year=datetime.now().year, hour=9, minute=0)
            return dt
        except Exception:
            pass

    m = re.match(r"^(\d{1,2})\s+([а-яё]+)(?:\s+(\d{4}))?(?:\s+в\s+(\d{1,2}:\d{2}))?$", expr.lower())
    if m:
        day = int(m.group(1))
        month_name = m.group(2)
        year = int(m.group(3)) if m.group(3) else datetime.now().year
        hhmm = m.group(4) or "09:00"
        months = {
            "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6,
            "июля": 7, "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12
        }
        if month_name not in months:
            return None
        hh, mm = map(int, hhmm.split(":"))
        try:
            return datetime(year, months[month_name], day, hh, mm)
        except Exception:
            return None

    return None


def parse_due_from_text(expr: str) -> Optional[datetime]:
    expr = expr.strip()

    rel = parse_relative_due(expr)
    if rel:
        return rel

    abs_dt = parse_absolute_due(expr)
    if abs_dt:
        return abs_dt

    if expr.lower().startswith("завтра"):
        m = re.search(r"завтра(?:\s+в\s+(\d{1,2}:\d{2}))?", expr.lower())
        if m:
            hhmm = m.group(1) or "09:00"
            hh, mm = map(int, hhmm.split(":"))
            base = datetime.now() + timedelta(days=1)
            return base.replace(hour=hh, minute=mm, second=0, microsecond=0)

    if expr.lower().startswith("сегодня"):
        m = re.search(r"сегодня(?:\s+в\s+(\d{1,2}:\d{2}))?", expr.lower())
        if m:
            hhmm = m.group(1) or "09:00"
            hh, mm = map(int, hhmm.split(":"))
            return datetime.now().replace(hour=hh, minute=mm, second=0, microsecond=0)

    return None


def parse_recurring_reminder(text: str) -> Optional[Dict[str, Any]]:
    t = text.strip().lower()

    m = re.match(r"^каждый\s+день\s+в\s+(\d{1,2}:\d{2})\s+(.+)$", t, re.I)
    if m:
        return {"rule": {"kind": "daily", "time": m.group(1)}, "text": m.group(2).strip()}

    m = re.match(r"^каждую?\s+неделю\s+в\s+(\d{1,2}:\d{2})\s+(.+)$", t, re.I)
    if m:
        return {"rule": {"kind": "weekly", "weekday": datetime.now().weekday(), "time": m.group(1)}, "text": m.group(2).strip()}

    for ru_day, wd in WEEKDAYS_RU.items():
        m = re.match(rf"^каждый\s+{ru_day}\s+в\s+(\d{{1,2}}:\d{{2}})\s+(.+)$", t, re.I)
        if m:
            return {"rule": {"kind": "weekly", "weekday": wd, "time": m.group(1)}, "text": m.group(2).strip()}

    return None


def parse_natural_reminder(text: str) -> Optional[Dict[str, Any]]:
    t = text.strip()

    m = re.match(r"^напомни(?:\s+мне)?\s+через\s+(\d+\s*[smhd])\s+(.+)$", t, re.I)
    if m:
        due = parse_relative_due(m.group(1))
        if due:
            return {"due": due, "text": m.group(2).strip(), "mode": "once"}

    m = re.match(r"^напомни(?:\s+мне)?\s+в\s+(.+?)\s+(.+)$", t, re.I)
    if m:
        due = parse_due_from_text(m.group(1))
        if due:
            return {"due": due, "text": m.group(2).strip(), "mode": "once"}

    recurring = parse_recurring_reminder(text)
    if recurring:
        due = next_recurring_due(recurring["rule"])
        if due:
            return {"due": due, "text": recurring["text"], "mode": "recurring", "rule": recurring["rule"]}

    return None


def parse_task_with_deadline(text: str) -> Optional[Dict[str, Any]]:
    m = re.match(r"^(.*?)(?:\s+до\s+(.+))$", text.strip(), re.I)
    if not m:
        return None

    task_text = m.group(1).strip()
    due_part = m.group(2).strip()
    due = parse_due_from_text(due_part)
    if not due:
        return None

    if len(task_text) < 3:
        return None

    return {"task_text": task_text, "due": due}


def parse_reminder_input(raw: str) -> Optional[Dict[str, Any]]:
    raw = raw.strip()
    parts = raw.split(None, 1)
    if len(parts) < 2:
        return None

    due = parse_due_from_text(parts[0])
    if due:
        return {"due": due, "text": parts[1].strip()}

    parts2 = raw.split(None, 2)
    if len(parts2) >= 3:
        due = parse_due_from_text(f"{parts2[0]} {parts2[1]}")
        if due:
            return {"due": due, "text": parts2[2].strip()}

    return None


def detect_translation_request(text: str) -> Optional[tuple]:
    t = text.strip()

    m = re.match(r"^(?:переведи|translate)\s+на\s+([^\:]+?)\s*[:\-]\s*(.+)$", t, re.I)
    if m:
        return m.group(1).strip(), m.group(2).strip()

    m = re.match(r"^(?:переведи|translate)\s+(.+?)\s+на\s+([а-яa-z\- ]+)$", t, re.I)
    if m:
        return m.group(2).strip(), m.group(1).strip()

    m = re.match(r"^как\s+это\s+на\s+([а-яa-z\- ]+)\??$", t, re.I)
    if m and len(t.split()) > 4:
        return m.group(1).strip(), ""

    return None


# =========================
# CONFIRMED ACTIONS
# =========================

async def execute_confirmed_action(action_type: str, payload: Dict[str, Any], user_id: int) -> str:
    if action_type == "delete_memory":
        memory_id = int(payload["memory_id"])
        await delete_memory(user_id, memory_id)
        return f"Память #{memory_id} удалена."

    if action_type == "reset_history":
        await clear_history(user_id)
        return "История очищена."

    if action_type == "set_setting":
        key = payload["key"]
        value = payload["value"]
        await set_setting(user_id, key, value)
        return "Настройка изменена."

    if action_type == "set_city":
        await set_setting(user_id, "default_city", payload["city"])
        return f"Город по умолчанию установлен: {payload['city']}"

    if action_type == "set_voice":
        await set_setting(user_id, "voice_mode", int(payload["value"]))
        return f"Voice mode: {'ON' if int(payload['value']) else 'OFF'}"

    if action_type == "set_model":
        await set_setting(user_id, "model_mode", payload["value"])
        return f"Model mode: {payload['value']}"

    if action_type == "set_briefing":
        await set_setting(user_id, "briefing_enabled", int(payload["value"]))
        val = int(payload["value"])
        return f"Briefing: {'ON' if val else 'OFF'}"

    return "Готово."


async def ask_confirmation(message: Message, title: str, action_type: str, payload: Dict[str, Any]) -> None:
    cid = await add_confirmation(message.from_user.id, action_type, payload)

    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Да", callback_data=f"confirm:{cid}"),
                InlineKeyboardButton(text="❌ Нет", callback_data=f"cancel:{cid}"),
            ]
        ]
    )
    await message.answer(f"{title}\n\nПодтвердить?", reply_markup=kb)


@router.callback_query(F.data.startswith("confirm:"))
async def cb_confirm_action(callback: CallbackQuery) -> None:
    user_id = callback.from_user.id
    cid = int(callback.data.split(":", 1)[1])
    row = await get_confirmation(cid)

    if not row or int(row["user_id"]) != user_id:
        await callback.answer("Подтверждение не найдено или уже устарело.", show_alert=True)
        return

    payload = json.loads(row["payload"])
    result = await execute_confirmed_action(row["action_type"], payload, user_id)
    await delete_confirmation(cid)
    await callback.answer("Готово")
    await callback.message.answer(result)


@router.callback_query(F.data.startswith("cancel:"))
async def cb_cancel_action(callback: CallbackQuery) -> None:
    cid = int(callback.data.split(":", 1)[1])
    await delete_confirmation(cid)
    await callback.answer("Отменено")
    await callback.message.answer("Действие отменено.")


# =========================
# NATURAL INTENTS
# =========================

async def handle_natural_language_intents(message: Message, text: str) -> Optional[str]:
    user_id = message.from_user.id
    t = normalize_text(text)

    tr = detect_translation_request(t)
    if tr:
        target_lang, source_text = tr
        if not source_text:
            return "Напиши сам текст, который нужно перевести, и язык."
        return await translate_text(source_text, target_lang)

    reminder = parse_natural_reminder(t)
    if reminder:
        due = reminder["due"].isoformat(timespec="minutes")
        text_to_remind = reminder["text"]

        if reminder.get("mode") == "recurring":
            rid = await add_recurring_reminder(
                user_id=user_id,
                chat_id=message.chat.id,
                text=text_to_remind,
                rule=reminder["rule"],
                next_due=reminder["due"],
            )
            return f"Периодическое напоминание сохранено. ID: {rid}\nПервый запуск: {due}"

        rid = await add_reminder(user_id, message.chat.id, text_to_remind, due)
        return f"Напоминание сохранено. ID: {rid}\nВремя: {due}"

    task_with_due = parse_task_with_deadline(t)
    if task_with_due:
        task_id = await add_task(user_id, task_with_due["task_text"], task_with_due["due"].isoformat(timespec="minutes"))
        due = task_with_due["due"]
        remind_at = due - timedelta(hours=1)

        if remind_at > datetime.now():
            await add_reminder(
                user_id,
                message.chat.id,
                f"Напоминание по задаче #{task_id}: {task_with_due['task_text']}",
                remind_at.isoformat(timespec="minutes"),
            )
        return f"Задача добавлена. ID: {task_id}\nДедлайн: {format_dt(due)}"

    if re.search(r"\b(удали|сотри|очисти)\b", t, re.I):
        if re.search(r"\bистори[юя]\b", t, re.I) or re.search(r"\bпамять\b", t, re.I):
            await ask_confirmation(
                message,
                "Ты хочешь очистить историю сообщений?",
                "reset_history",
                {},
            )
            return ""

        m = re.search(r"(?:удали|сотри)\s+(?:память|memory)\s*(\d+)", t, re.I)
        if m:
            await ask_confirmation(
                message,
                f"Удалить память #{m.group(1)}?",
                "delete_memory",
                {"memory_id": int(m.group(1))},
            )
            return ""

    if re.search(r"\b(выключи|включи|поменяй|измени|установи)\b", t, re.I):
        if re.search(r"\bголос\b", t, re.I):
            new_value = 0 if re.search(r"\bвыключи\b", t, re.I) else 1
            await ask_confirmation(
                message,
                f"Изменить voice mode на {'ON' if new_value else 'OFF'}?",
                "set_voice",
                {"value": new_value},
            )
            return ""

        if re.search(r"\bбрифинг\b", t, re.I):
            new_value = 0 if re.search(r"\bвыключи\b", t, re.I) else 1
            await ask_confirmation(
                message,
                f"Изменить briefing на {'ON' if new_value else 'OFF'}?",
                "set_briefing",
                {"value": new_value},
            )
            return ""

        if re.search(r"\bмодель\b", t, re.I):
            m = re.search(r"\b(fast|smart|vision)\b", t, re.I)
            if m:
                await ask_confirmation(
                    message,
                    f"Изменить model mode на {m.group(1).lower()}?",
                    "set_model",
                    {"value": m.group(1).lower()},
                )
                return ""

        if re.search(r"\bгород\b", t, re.I):
            m = re.search(r"(?:на|в)\s+([А-Яа-яЁёA-Za-z\- ]+)$", t)
            if m:
                city = normalize_text(m.group(1))
                await ask_confirmation(
                    message,
                    f"Изменить город по умолчанию на «{city}»?",
                    "set_city",
                    {"city": city},
                )
                return ""

    return None


# =========================
# CORE REPLY
# =========================

async def generate_jarvis_reply(user_id: int, text: str, save_user_message: bool = True) -> Optional[str]:
    await ensure_user(user_id)
    text = normalize_text(text)
    if not text:
        return None

    if save_user_message:
        await save_message(user_id, "user", text)

    if text.lower().startswith("запомни "):
        content = text.split(" ", 1)[1].strip()
        if content:
            mid = await add_memory(user_id, content)
            answer = f"Запомнил. ID памяти: {mid}"
            await save_message(user_id, "assistant", answer)
            return answer
        return None

    if text.lower().startswith("заметка "):
        content = text.split(" ", 1)[1].strip()
        if content:
            nid = await add_note(user_id, title=content[:40], content=content)
            answer = f"Заметка сохранена. ID: {nid}"
            await save_message(user_id, "assistant", answer)
            return answer
        return None

    settings = await get_settings(user_id)
    system_prompt = await build_system_prompt(user_id, text)
    history = await get_recent_messages(user_id, limit=6)
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

    await save_message(user_id, "user", text)

    intent_reply = await handle_natural_language_intents(message, text)
    if intent_reply is not None:
        if intent_reply.strip():
            await send_text_and_optional_voice(message, intent_reply, settings)
        return

    await auto_learn_from_text(user_id, text)

    answer = await generate_jarvis_reply(user_id, text, save_user_message=False)
    if answer is None:
        return

    # Не показываем proactive advice если пришла ошибка от Groq
    if not answer.startswith("Ошибка Groq") and not answer.startswith("Groq API временно"):
        proactive = await build_proactive_advice(user_id)
        if proactive:
            answer = f"{answer}\n\n{proactive}"

    await send_text_and_optional_voice(message, answer, settings)


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
        "• обычный диалог с памятью\n"
        "• авто-сохранение важных фактов из разговора\n"
        "• перевод без слеш-команд\n"
        "• напоминания из обычного текста\n"
        "• периодические напоминания\n"
        "• задачи с дедлайнами и авто-напоминанием за час\n"
        "• поиск по памяти, задачам, заметкам и документам\n"
        "• разбор файлов и построение своей библиотеки\n"
        "• трекинг сна, курения и самочувствия\n"
        "• анализ медицинских файлов и последующие советы\n"
        "• голосовой ввод/ответ\n"
        "• фото и документы\n\n"
        "Важно:\n"
        "• удаление и любые изменения настроек — только с подтверждением\n\n"
        "Пиши просто как обычно, без лишних команд."
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
    await ask_confirmation(
        message,
        "Очистить историю сообщений?",
        "reset_history",
        {},
    )


@router.message(Command("voice"))
async def cmd_voice(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip().lower()
    if arg not in {"on", "off"}:
        await message.answer("Используй: /voice on или /voice off")
        return

    new_value = 1 if arg == "on" else 0
    await ask_confirmation(
        message,
        f"Изменить voice mode на {'ON' if new_value else 'OFF'}?",
        "set_voice",
        {"value": new_value},
    )


@router.message(Command("model"))
async def cmd_model(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip().lower()
    if arg not in {"fast", "smart", "vision"}:
        await message.answer("Используй: /model fast, /model smart или /model vision")
        return

    await ask_confirmation(
        message,
        f"Изменить model mode на {arg}?",
        "set_model",
        {"value": arg},
    )


@router.message(Command("briefing"))
async def cmd_briefing(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip().lower()
    if arg not in {"on", "off"}:
        await message.answer("Используй: /briefing on или /briefing off")
        return

    await ask_confirmation(
        message,
        f"Изменить briefing на {arg.upper()}?",
        "set_briefing",
        {"value": 1 if arg == "on" else 0},
    )


@router.message(Command("brief"))
async def cmd_brief(message: Message) -> None:
    user_id = message.from_user.id
    briefing = await build_briefing(user_id)
    await message.answer(briefing)


@router.message(Command("city"))
async def cmd_city(message: Message, command: CommandObject) -> None:
    city = (command.args or "").strip()
    if not city:
        await message.answer("Используй: /city Москва")
        return

    await ask_confirmation(
        message,
        f"Изменить город по умолчанию на «{city}»?",
        "set_city",
        {"city": city},
    )


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
    arg = (command.args or "").strip()
    if not arg.isdigit():
        await message.answer("Используй: /forget ID")
        return

    await ask_confirmation(
        message,
        f"Удалить память #{int(arg)}?",
        "delete_memory",
        {"memory_id": int(arg)},
    )


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
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /search memory запрос\nили: /search tasks запрос")
        return

    scope = "all"
    query = raw

    low = raw.lower()
    for prefix, sc in (
        ("memory ", "memory"),
        ("mem ", "memory"),
        ("tasks ", "tasks"),
        ("task ", "tasks"),
        ("docs ", "docs"),
        ("doc ", "docs"),
        ("notes ", "notes"),
        ("note ", "notes"),
    ):
        if low.startswith(prefix):
            scope = sc
            query = raw[len(prefix):].strip()
            break

    results = await search_knowledge(query, user_id, top_k=5, scope=scope)
    if not results:
        await message.answer("Ничего не нашёл.")
        return

    lines = [f"Результаты поиска ({scope}):"]
    for r in results:
        lines.append(f"• [{r['source']} #{r['id']}] {r['text'][:220]} (score={r['score']:.3f})")
    await message.answer("\n".join(lines))


@router.message(Command("ask"))
async def cmd_ask(message: Message, command: CommandObject) -> None:
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

    await ask_confirmation(
        callback.message,
        f"Изменить voice mode на {'ON' if new_value else 'OFF'}?",
        "set_voice",
        {"value": new_value},
    )
    await callback.answer("Запрос на подтверждение отправлен")


@router.callback_query(F.data == "jarvis_briefing")
async def cb_briefing(callback: CallbackQuery) -> None:
    user_id = callback.from_user.id
    settings = await get_settings(user_id)
    new_value = 0 if int(settings.get("briefing_enabled", 0)) else 1

    await ask_confirmation(
        callback.message,
        f"Изменить briefing на {'ON' if new_value else 'OFF'}?",
        "set_briefing",
        {"value": new_value},
    )
    await callback.answer("Запрос на подтверждение отправлен")


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
        await add_document(
            user_id=message.from_user.id,
            filename=filename,
            content=text,
            summary=summary,
            doc_type=ext,
        )

        await auto_learn_from_text(message.from_user.id, text, source="document")

        if re.search(r"(медицин|анализ|врач|заключени|диагноз|обследован)", filename + " " + text, re.I):
            await add_memory(
                message.from_user.id,
                f"Загружен медицинский документ: {filename}. Кратко: {summary[:400]}",
            )

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


async def recurring_reminders_job() -> None:
    reminders = await get_due_recurring_reminders(limit=100)
    if not reminders:
        return

    for r in reminders:
        try:
            rule = json.loads(r["rule"])
            await bot.send_message(
                r["chat_id"],
                f"⏰ Напоминание #{r['id']}:\n{r['text']}\n\nЭто периодическое напоминание.",
            )
            next_due = next_recurring_due(rule, datetime.now() + timedelta(minutes=1))
            if next_due:
                await update_recurring_reminder_next_due(r["id"], next_due)
        except Exception as e:
            logger.warning("Failed recurring reminder #%s: %s", r["id"], e)


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
# HTTP API
# =========================

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
        answer = await generate_jarvis_reply(user_id, text, save_user_message=True)
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
    scheduler.add_job(recurring_reminders_job, "interval", seconds=30, id="recurring_reminders_job", replace_existing=True)
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
