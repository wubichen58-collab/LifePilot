import asyncio
import base64
import json
import logging
import os
import re
import sqlite3
import tempfile
import math
import statistics
from collections import Counter
from datetime import datetime, timedelta
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
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
SERPER_API_KEY = os.environ.get("SERPER_API_KEY")  # для веб-поиска
NEWSAPI_KEY = os.environ.get("NEWSAPI_KEY")  # для новостей

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
# SELF-REPAIR LOG
# =========================

SELF_REPAIR_LOG: List[Dict[str, Any]] = []

def log_self_repair(module: str, issue: str, fix: str) -> None:
    entry = {
        "time": now_str(),
        "module": module,
        "issue": issue,
        "fix": fix,
    }
    SELF_REPAIR_LOG.append(entry)
    logger.warning("[SELF-REPAIR] %s | %s | %s", module, issue, fix)


# =========================
# DB
# =========================

def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        cur.execute("""
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
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                role TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS memories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                text TEXT NOT NULL,
                status TEXT DEFAULT 'open',
                priority INTEGER DEFAULT 2,
                due_at TEXT DEFAULT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS reminders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                chat_id INTEGER NOT NULL,
                text TEXT NOT NULL,
                due_at TEXT NOT NULL,
                sent INTEGER DEFAULT 0,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS confirmations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                action_type TEXT NOT NULL,
                payload TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                expires_at TEXT DEFAULT NULL
            )
        """)

        cur.execute("""
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
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                filename TEXT NOT NULL,
                content TEXT NOT NULL,
                summary TEXT NOT NULL,
                doc_type TEXT DEFAULT 'file',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
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
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS threat_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                threat_type TEXT NOT NULL,
                description TEXT NOT NULL,
                severity INTEGER DEFAULT 1,
                resolved INTEGER DEFAULT 0,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS analytics_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                event_type TEXT NOT NULL,
                data_json TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS finance_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                record_type TEXT NOT NULL,
                amount REAL NOT NULL,
                category TEXT DEFAULT NULL,
                note TEXT DEFAULT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS behavior_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                pattern TEXT NOT NULL,
                count INTEGER DEFAULT 1,
                last_seen TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

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
        "user_id": user_id, "voice_mode": 0, "model_mode": "smart",
        "briefing_enabled": 0, "briefing_time": "09:00",
        "default_city": None, "last_brief_date": None,
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
        "SELECT role, content FROM messages WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, limit),
    )
    rows.reverse()
    return [{"role": r["role"], "content": r["content"]} for r in rows]


async def clear_history(user_id: int) -> None:
    await db_execute("DELETE FROM messages WHERE user_id = ?", (user_id,))


async def add_memory(user_id: int, content: str) -> int:
    return await db_execute(
        "INSERT INTO memories (user_id, content) VALUES (?, ?)", (user_id, content),
    )


async def get_memories(user_id: int, limit: int = 20) -> List[Dict[str, Any]]:
    return await db_query(
        "SELECT id, content, created_at FROM memories WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, limit),
    )


async def delete_memory(user_id: int, memory_id: int) -> None:
    await db_execute("DELETE FROM memories WHERE user_id = ? AND id = ?", (user_id, memory_id))


async def add_note(user_id: int, title: str, content: str) -> int:
    return await db_execute(
        "INSERT INTO notes (user_id, title, content) VALUES (?, ?, ?)", (user_id, title, content),
    )


async def get_notes(user_id: int, limit: int = 20) -> List[Dict[str, Any]]:
    return await db_query(
        "SELECT id, title, content, created_at FROM notes WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, limit),
    )


async def add_task(user_id: int, text: str, due_at: Optional[str] = None, priority: int = 2) -> int:
    return await db_execute(
        "INSERT INTO tasks (user_id, text, due_at, priority) VALUES (?, ?, ?, ?)",
        (user_id, text, due_at, priority),
    )


async def get_tasks(user_id: int, status: str = "open", limit: int = 20) -> List[Dict[str, Any]]:
    return await db_query(
        "SELECT id, text, status, due_at, priority, created_at FROM tasks WHERE user_id = ? AND status = ? ORDER BY priority ASC, id DESC LIMIT ?",
        (user_id, status, limit),
    )


async def mark_task_done(user_id: int, task_id: int) -> bool:
    affected = await db_exec_rowcount(
        "UPDATE tasks SET status = 'done' WHERE user_id = ? AND id = ?", (user_id, task_id),
    )
    return affected > 0


async def add_reminder(user_id: int, chat_id: int, text: str, due_at: str) -> int:
    return await db_execute(
        "INSERT INTO reminders (user_id, chat_id, text, due_at) VALUES (?, ?, ?, ?)",
        (user_id, chat_id, text, due_at),
    )


async def get_due_reminders(limit: int = 100) -> List[Dict[str, Any]]:
    now_iso = datetime.now().isoformat(timespec="minutes")
    return await db_query(
        "SELECT id, user_id, chat_id, text, due_at FROM reminders WHERE sent = 0 AND due_at <= ? ORDER BY due_at ASC LIMIT ?",
        (now_iso, limit),
    )


async def mark_reminder_sent(reminder_id: int) -> None:
    await db_exec_rowcount("UPDATE reminders SET sent = 1 WHERE id = ?", (reminder_id,))


async def get_due_reminders_for_user(user_id: int) -> List[Dict[str, Any]]:
    now_iso = datetime.now().isoformat(timespec="minutes")
    return await db_query(
        "SELECT id, user_id, chat_id, text, due_at FROM reminders WHERE sent = 0 AND user_id = ? AND due_at <= ? ORDER BY due_at ASC",
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
        "INSERT INTO documents (user_id, filename, content, summary, doc_type) VALUES (?, ?, ?, ?, ?)",
        (user_id, filename, content, summary, doc_type),
    )


async def get_documents(user_id: int, limit: int = 20) -> List[Dict[str, Any]]:
    return await db_query(
        "SELECT id, filename, content, summary, doc_type, created_at FROM documents WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, limit),
    )


async def add_health_metric(
    user_id: int, metric: str, value: Optional[float] = None,
    text_value: Optional[str] = None, note: Optional[str] = None,
    meta: Optional[Dict[str, Any]] = None,
) -> int:
    return await db_execute(
        "INSERT INTO health_metrics (user_id, metric, value, text_value, note, meta_json) VALUES (?, ?, ?, ?, ?, ?)",
        (user_id, metric, value, text_value, note, json.dumps(meta, ensure_ascii=False) if meta else None),
    )


async def get_health_metrics(user_id: int, limit: int = 50) -> List[Dict[str, Any]]:
    return await db_query(
        "SELECT id, metric, value, text_value, note, meta_json, created_at FROM health_metrics WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, limit),
    )


async def log_threat(user_id: int, threat_type: str, description: str, severity: int = 1) -> int:
    return await db_execute(
        "INSERT INTO threat_log (user_id, threat_type, description, severity) VALUES (?, ?, ?, ?)",
        (user_id, threat_type, description, severity),
    )


async def log_analytics_event(user_id: int, event_type: str, data: Dict[str, Any]) -> None:
    await db_execute(
        "INSERT INTO analytics_events (user_id, event_type, data_json) VALUES (?, ?, ?)",
        (user_id, event_type, json.dumps(data, ensure_ascii=False)),
    )


async def add_finance_record(user_id: int, record_type: str, amount: float, category: str = "", note: str = "") -> int:
    return await db_execute(
        "INSERT INTO finance_records (user_id, record_type, amount, category, note) VALUES (?, ?, ?, ?, ?)",
        (user_id, record_type, amount, category, note),
    )


async def get_finance_records(user_id: int, limit: int = 50) -> List[Dict[str, Any]]:
    return await db_query(
        "SELECT * FROM finance_records WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, limit),
    )


async def update_behavior_pattern(user_id: int, pattern: str) -> None:
    existing = await db_fetch_one(
        "SELECT id, count FROM behavior_log WHERE user_id = ? AND pattern = ?", (user_id, pattern)
    )
    if existing:
        await db_execute(
            "UPDATE behavior_log SET count = count + 1, last_seen = ? WHERE id = ?",
            (now_str(), existing["id"]),
        )
    else:
        await db_execute(
            "INSERT INTO behavior_log (user_id, pattern, count, last_seen) VALUES (?, ?, 1, ?)",
            (user_id, pattern, now_str()),
        )


async def get_behavior_patterns(user_id: int) -> List[Dict[str, Any]]:
    return await db_query(
        "SELECT pattern, count, last_seen FROM behavior_log WHERE user_id = ? ORDER BY count DESC LIMIT 20",
        (user_id,),
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


async def add_recurring_reminder(user_id: int, chat_id: int, text: str, rule: Dict[str, Any], next_due: datetime) -> int:
    return await db_execute(
        "INSERT INTO recurring_reminders (user_id, chat_id, text, rule, next_due) VALUES (?, ?, ?, ?, ?)",
        (user_id, chat_id, text, json.dumps(rule, ensure_ascii=False), next_due.isoformat(timespec="minutes")),
    )


async def get_due_recurring_reminders(limit: int = 100) -> List[Dict[str, Any]]:
    now_iso = datetime.now().isoformat(timespec="minutes")
    return await db_query(
        "SELECT id, user_id, chat_id, text, rule, next_due, sent_count FROM recurring_reminders WHERE active = 1 AND next_due <= ? ORDER BY next_due ASC LIMIT ?",
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
    "понедельник": 0, "вторник": 1, "среда": 2, "четверг": 3,
    "пятница": 4, "суббота": 5, "воскресенье": 6,
}

PRIORITY_MAP = {"высокий": 1, "высок": 1, "срочно": 1, "средний": 2, "средн": 2, "низкий": 3, "низк": 3}


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
                InlineKeyboardButton(text="🛡 Угрозы", callback_data="jarvis_threats"),
                InlineKeyboardButton(text="📊 Аналитика", callback_data="jarvis_analytics"),
            ],
            [
                InlineKeyboardButton(text="💰 Финансы", callback_data="jarvis_finance"),
                InlineKeyboardButton(text="🔧 Статус", callback_data="jarvis_status"),
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
    payload = {"model": model, "messages": messages, "temperature": temperature}

    for attempt in range(3):
        try:
            async with session.post(GROQ_CHAT_URL, headers=headers, json=payload) as resp:
                data = await resp.json(content_type=None)
                if resp.status == 429:
                    wait = 20 * (attempt + 1)
                    logger.warning("Rate limit, жду %s сек (попытка %s)...", wait, attempt + 1)
                    await asyncio.sleep(wait)
                    continue
                if resp.status >= 400:
                    logger.error("Groq error %s: %s", resp.status, data)
                    log_self_repair("groq_chat", f"HTTP {resp.status}", "logged, retry")
                    return f"Ошибка Groq API: {data}"
                return data["choices"][0]["message"]["content"].strip()
        except Exception as e:
            log_self_repair("groq_chat", str(e), "exception caught, retry")
            await asyncio.sleep(5)
            continue

    return "Groq API временно недоступен, попробуй чуть позже."


async def groq_transcribe(audio_bytes: bytes, filename: str = "voice.ogg") -> Optional[str]:
    if not GROQ_API_KEY:
        return None
    session = await get_http_session()
    form = aiohttp.FormData()
    form.add_field("file", audio_bytes, filename=filename, content_type="audio/ogg")
    form.add_field("model", GROQ_MODEL_ASR)
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    try:
        async with session.post(GROQ_TRANSCRIBE_URL, headers=headers, data=form) as resp:
            data = await resp.json(content_type=None)
            if resp.status >= 400:
                logger.error("Transcription error %s: %s", resp.status, data)
                return None
            return (data.get("text") or "").strip() or None
    except Exception as e:
        log_self_repair("groq_transcribe", str(e), "exception caught")
        return None


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
        log_self_repair("tts_to_mp3", str(e), "skipped TTS")
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
# SEARCH (семантический)
# =========================

def semantic_search_sync(query: str, items: List[Dict[str, Any]], top_k: int = 5) -> List[Dict[str, Any]]:
    query_norm = query.lower().strip()
    tokens = [t for t in query_norm.split() if len(t) > 2]
    if not tokens or not items:
        return []
    scored = []
    for item in items:
        text_lower = item["text"].lower()
        score = sum(1.0 for t in tokens if t in text_lower)
        for i in range(len(tokens) - 1):
            if f"{tokens[i]} {tokens[i+1]}" in text_lower:
                score += 1.5
        if score > 0:
            scored.append({**item, "score": round(score / (1 + len(text_lower) / 500), 3)})
    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored[:top_k]


async def search_knowledge(query: str, user_id: int, top_k: int = 3, scope: str = "all") -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    if scope in {"all", "memory"}:
        for row in await get_memories(user_id, limit=120):
            items.append({"source": "memory", "id": row["id"], "text": row["content"]})
    if scope in {"all", "notes"}:
        for row in await get_notes(user_id, limit=120):
            items.append({"source": "note", "id": row["id"], "text": f"{row['title']}: {row['content']}"})
    if scope in {"all", "tasks"}:
        for row in await get_tasks(user_id, status="open", limit=120):
            due = f" | due: {row['due_at']}" if row.get("due_at") else ""
            items.append({"source": "task", "id": row["id"], "text": f"{row['text']} | status: {row['status']}{due}"})
    if scope in {"all", "docs"}:
        for row in await get_documents(user_id, limit=80):
            items.append({"source": "doc", "id": row["id"], "text": f"{row['filename']}: {row['summary']} | {row['content'][:3000]}"})
    if not items:
        return []
    return await asyncio.to_thread(semantic_search_sync, query, items, top_k)


# =========================
# WEB SEARCH
# =========================

async def web_search(query: str, num: int = 5) -> str:
    """Поиск через Serper.dev (бесплатный план)"""
    if not SERPER_API_KEY:
        return await _fallback_web_search(query)
    session = await get_http_session()
    try:
        async with session.post(
            "https://google.serper.dev/search",
            headers={"X-API-KEY": SERPER_API_KEY, "Content-Type": "application/json"},
            json={"q": query, "num": num, "hl": "ru"},
        ) as resp:
            if resp.status != 200:
                return await _fallback_web_search(query)
            data = await resp.json(content_type=None)
            results = data.get("organic", [])
            if not results:
                return "Поиск не дал результатов."
            lines = [f"🔍 Результаты поиска: «{query}»"]
            for r in results[:5]:
                lines.append(f"• {r.get('title', '—')}\n  {r.get('snippet', '')}\n  🔗 {r.get('link', '')}")
            return "\n\n".join(lines)
    except Exception as e:
        log_self_repair("web_search", str(e), "fallback to LLM")
        return await _fallback_web_search(query)


async def _fallback_web_search(query: str) -> str:
    """Если нет Serper — используем Groq как суррогат поиска"""
    return await groq_chat(
        [
            {"role": "system", "content": "Ты поисковый агент. Дай краткий, фактический ответ на запрос. Если не знаешь точно — скажи об этом."},
            {"role": "user", "content": f"Найди информацию: {query}"},
        ],
        model=GROQ_MODEL_FAST,
        temperature=0.1,
    )


async def fetch_news(query: str = "", category: str = "general") -> str:
    """Получение новостей через NewsAPI"""
    if not NEWSAPI_KEY:
        return await groq_chat(
            [{"role": "system", "content": "Дай краткий обзор последних новостей по теме."}, {"role": "user", "content": f"Новости: {query or category}"}],
            model=GROQ_MODEL_FAST, temperature=0.2,
        )
    session = await get_http_session()
    try:
        params = {"apiKey": NEWSAPI_KEY, "language": "ru", "pageSize": 5}
        if query:
            params["q"] = query
        else:
            params["category"] = category
            params["country"] = "ru"
        url = "https://newsapi.org/v2/top-headlines" if not query else "https://newsapi.org/v2/everything"
        async with session.get(url, params=params) as resp:
            if resp.status != 200:
                return "Не удалось получить новости."
            data = await resp.json(content_type=None)
            articles = data.get("articles", [])
            if not articles:
                return "Новостей по теме не найдено."
            lines = [f"📰 Новости: {query or category}"]
            for a in articles[:5]:
                lines.append(f"• {a.get('title', '—')}\n  {a.get('description', '')}\n  🔗 {a.get('url', '')}")
            return "\n\n".join(lines)
    except Exception as e:
        log_self_repair("fetch_news", str(e), "skipped")
        return "Ошибка при получении новостей."


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
            return None
        locs = data.get("location") or []
        return locs[0] if locs else None


async def qweather_now(location_id: str) -> Optional[Dict[str, Any]]:
    if not QWEATHER_KEY:
        return None
    session = await get_http_session()
    url = f"https://devapi.qweather.com/v7/weather/now?location={location_id}&key={QWEATHER_KEY}"
    async with session.get(url) as resp:
        data = await resp.json(content_type=None)
        if resp.status >= 400:
            return None
        return data.get("now")


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
    system = "Ты кратко и точно суммируешь документы по-русски. Дай 5-10 буллетов, потом 1 короткий вывод. Не выдумывай факты."
    user = f"Название: {title}\n\nТекст:\n{text}"
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": user}],
        model=GROQ_MODEL_FAST, temperature=0.2,
    )


def detect_file_type(filename: str) -> str:
    return filename.lower().rsplit(".", 1)[-1] if "." in filename else ""


# =========================
# IMAGE ANALYSIS
# =========================

async def analyze_image_with_groq(image_bytes: bytes, caption: str = "", mode: str = "general") -> str:
    if not GROQ_API_KEY:
        return "GROQ_API_KEY не задан."
    session = await get_http_session()
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    data_url = f"data:image/jpeg;base64,{b64}"

    system_prompts = {
        "general": "Ты анализируешь изображение. Отвечай по-русски, кратко, но полезно. Если есть текст — извлеки его.",
        "chemical": (
            "Ты проводишь поверхностный химический анализ изображения. "
            "Определи: видимые вещества, цвет, текстуру, возможный состав, агрегатное состояние, "
            "признаки реакций (коррозия, окисление, кристаллизация). "
            "Дай научное объяснение наблюдаемых явлений. Отвечай структурированно."
        ),
        "threat": (
            "Ты анализируешь изображение на наличие угроз и опасностей. "
            "Определи: потенциальные угрозы, опасные объекты, подозрительные элементы, "
            "уязвимые места. Дай оценку уровня угрозы от 1 до 10."
        ),
        "face": (
            "Ты анализируешь фотографию человека. "
            "Опиши: возраст (приблизительно), эмоциональное состояние, поведенческие признаки, "
            "уровень стресса, общее впечатление. НЕ называй имён — только поведенческий анализ."
        ),
        "construction": (
            "Ты анализируешь конструкцию на изображении. "
            "Определи: тип конструкции, материалы, видимые дефекты или слабые места, "
            "структурную целостность, рекомендации."
        ),
        "medical": (
            "Ты анализируешь медицинское изображение или симптомы. "
            "Дай описательный анализ, возможные интерпретации. "
            "ВАЖНО: это не медицинский диагноз — рекомендуй обратиться к врачу."
        ),
    }

    system = system_prompts.get(mode, system_prompts["general"])
    default_captions = {
        "chemical": "Проведи химический поверхностный анализ изображения.",
        "threat": "Проанализируй на угрозы и опасности.",
        "face": "Проведи поведенческий анализ человека на фото.",
        "construction": "Проанализируй конструкцию.",
        "medical": "Проанализируй медицинское изображение.",
    }
    user_text = caption or default_captions.get(mode, "Опиши изображение.")

    payload = {
        "model": GROQ_MODEL_VISION,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": [
                {"type": "text", "text": user_text},
                {"type": "image_url", "image_url": {"url": data_url}},
            ]},
        ],
        "temperature": 0.2,
    }
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    try:
        async with session.post(GROQ_CHAT_URL, headers=headers, json=payload) as resp:
            data = await resp.json(content_type=None)
            if resp.status >= 400:
                return "Не удалось проанализировать изображение."
            return data["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_self_repair("analyze_image", str(e), "skipped")
        return "Ошибка при анализе изображения."


# =========================
# THREAT ANALYSIS
# =========================

async def analyze_threat(description: str, user_id: int) -> str:
    """Анализ угроз и определение слабых мест"""
    system = (
        "Ты аналитик безопасности. Твоя задача — провести анализ угрозы.\n"
        "Структура ответа:\n"
        "1. ТИП УГРОЗЫ\n"
        "2. УРОВЕНЬ ОПАСНОСТИ (1-10)\n"
        "3. СЛАБЫЕ МЕСТА ПРОТИВНИКА / ИСТОЧНИКА УГРОЗЫ\n"
        "4. ВЕРОЯТНОСТЬ РЕАЛИЗАЦИИ (%)\n"
        "5. ТАКТИЧЕСКИЕ РЕКОМЕНДАЦИИ\n"
        "6. ВЫВОД\n"
        "Будь конкретен, логичен, без воды."
    )
    result = await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": description}],
        model=GROQ_MODEL_TEXT, temperature=0.3,
    )

    severity = 5
    m = re.search(r"УРОВЕНЬ ОПАСНОСТИ[:\s]+(\d+)", result, re.I)
    if m:
        severity = min(int(m.group(1)), 10)

    await log_threat(user_id, "manual_analysis", description[:500], severity)
    await log_analytics_event(user_id, "threat_analysis", {"severity": severity})
    return result


async def calculate_win_probability(my_strengths: str, opponent_weaknesses: str) -> str:
    """Расчёт вероятности победы"""
    system = (
        "Ты военный аналитик и стратег. Рассчитай вероятность победы.\n"
        "Используй: SWOT-анализ, оценку ресурсов, тактические факторы.\n"
        "Формат:\n"
        "• Мои сильные стороны: оценка\n"
        "• Слабости противника: оценка\n"
        "• Расчёт вероятности победы: X%\n"
        "• Ключевые факторы успеха\n"
        "• Тактические рекомендации\n"
        "• Риски и как их минимизировать"
    )
    return await groq_chat(
        [
            {"role": "system", "content": system},
            {"role": "user", "content": f"Мои сильные стороны: {my_strengths}\n\nСлабости противника: {opponent_weaknesses}"},
        ],
        model=GROQ_MODEL_TEXT, temperature=0.3,
    )


async def tactical_recommendations(situation: str) -> str:
    """Тактические рекомендации"""
    system = (
        "Ты тактический советник. Анализируй ситуацию и давай конкретные, применимые рекомендации.\n"
        "Структура: Оценка ситуации → Варианты действий → Приоритетный план → Запасные варианты → Риски."
    )
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": situation}],
        model=GROQ_MODEL_TEXT, temperature=0.4,
    )


# =========================
# SCIENTIFIC CALCULATIONS
# =========================

def calculate_physics(formula: str, variables: Dict[str, float]) -> Dict[str, Any]:
    """Базовые физические расчёты"""
    results = {}
    try:
        formula_clean = formula.lower().strip()

        # F = ma
        if "f = ma" in formula_clean or "сила" in formula_clean:
            if "m" in variables and "a" in variables:
                results["F (сила, Н)"] = variables["m"] * variables["a"]

        # E = mc²
        if "e = mc" in formula_clean or "энергия покоя" in formula_clean:
            if "m" in variables:
                c = 3e8
                results["E (энергия, Дж)"] = variables["m"] * c ** 2

        # v = s/t
        if "скорость" in formula_clean or "v = s/t" in formula_clean:
            if "s" in variables and "t" in variables and variables["t"] != 0:
                results["v (скорость, м/с)"] = variables["s"] / variables["t"]

        # E_kinetic = mv²/2
        if "кинетическая" in formula_clean or "ek" in formula_clean:
            if "m" in variables and "v" in variables:
                results["Ek (кинетическая энергия, Дж)"] = 0.5 * variables["m"] * variables["v"] ** 2

        # P = F/S давление
        if "давление" in formula_clean or "p = f/s" in formula_clean:
            if "f" in variables and "s" in variables and variables["s"] != 0:
                results["P (давление, Па)"] = variables["f"] / variables["s"]

        # Закон Ома
        if "ом" in formula_clean or "u = ir" in formula_clean:
            if "u" in variables and "r" in variables and variables["r"] != 0:
                results["I (ток, А)"] = variables["u"] / variables["r"]
            if "i" in variables and "r" in variables:
                results["U (напряжение, В)"] = variables["i"] * variables["r"]

        if not results:
            results["info"] = "Формула не распознана. Укажи конкретные переменные."

    except Exception as e:
        results["error"] = str(e)
    return results


def calculate_statistics(data: List[float]) -> Dict[str, Any]:
    """Статистический анализ данных"""
    if not data:
        return {"error": "Нет данных"}
    try:
        result = {
            "n": len(data),
            "сумма": round(sum(data), 4),
            "среднее": round(statistics.mean(data), 4),
            "медиана": round(statistics.median(data), 4),
            "мин": round(min(data), 4),
            "макс": round(max(data), 4),
            "размах": round(max(data) - min(data), 4),
        }
        if len(data) >= 2:
            result["дисперсия"] = round(statistics.variance(data), 4)
            result["стд_отклонение"] = round(statistics.stdev(data), 4)
        if len(data) >= 3:
            mean = statistics.mean(data)
            std = statistics.stdev(data)
            if std > 0:
                skewness = sum((x - mean) ** 3 for x in data) / (len(data) * std ** 3)
                result["асимметрия"] = round(skewness, 4)
        counter = Counter(data)
        mode_val = counter.most_common(1)
        if mode_val and mode_val[0][1] > 1:
            result["мода"] = mode_val[0][0]
        return result
    except Exception as e:
        return {"error": str(e)}


def calculate_math(expression: str) -> str:
    """Безопасный математический калькулятор"""
    safe_globals = {
        "__builtins__": {},
        "math": math,
        "sqrt": math.sqrt,
        "sin": math.sin,
        "cos": math.cos,
        "tan": math.tan,
        "log": math.log,
        "log10": math.log10,
        "exp": math.exp,
        "pi": math.pi,
        "e": math.e,
        "abs": abs,
        "round": round,
        "pow": pow,
        "factorial": math.factorial,
    }
    try:
        cleaned = re.sub(r"[^0-9+\-*/().%^√mathsincotaelgqrb ,\s]", "", expression.lower())
        cleaned = cleaned.replace("^", "**").replace("√", "sqrt(")
        result = eval(cleaned, safe_globals)
        return f"Результат: {result}"
    except Exception as e:
        return f"Ошибка вычисления: {e}"


async def scientific_analysis(query: str) -> str:
    """Научный анализ через LLM"""
    system = (
        "Ты научный аналитик. Отвечай точно, используй формулы, единицы измерения, научную терминологию. "
        "Структура: Теория → Расчёты/Модель → Выводы → Применение."
    )
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": query}],
        model=GROQ_MODEL_TEXT, temperature=0.2,
    )


async def chemical_model(substance: str) -> str:
    """Химический анализ вещества"""
    system = (
        "Ты химик-аналитик. Предоставь информацию о веществе:\n"
        "• Химическая формула и структура\n"
        "• Физические свойства (температура кипения/плавления, плотность)\n"
        "• Химические свойства и реакции\n"
        "• Применение\n"
        "• Безопасность и токсичность\n"
        "• Методы обнаружения\n"
        "Используй научные данные."
    )
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": f"Анализ вещества: {substance}"}],
        model=GROQ_MODEL_TEXT, temperature=0.1,
    )


async def physical_model(phenomenon: str) -> str:
    """Физическая модель явления"""
    system = (
        "Ты физик-теоретик. Опиши физическую модель:\n"
        "• Уравнения и законы\n"
        "• Граничные условия\n"
        "• Численные параметры\n"
        "• Графическое описание\n"
        "• Практическое применение"
    )
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": f"Физическая модель: {phenomenon}"}],
        model=GROQ_MODEL_TEXT, temperature=0.2,
    )


# =========================
# MEDICAL & HEALTH
# =========================

async def analyze_virus(name: str) -> str:
    """Анализ вируса/патогена"""
    system = (
        "Ты вирусолог и эпидемиолог. Предоставь научную информацию:\n"
        "• Классификация и строение\n"
        "• Механизм заражения и репликации\n"
        "• Симптомы и течение болезни\n"
        "• Методы диагностики\n"
        "• Лечение и профилактика\n"
        "• Эпидемиологические данные\n"
        "• Угроза и уровень опасности\n"
        "Используй актуальные научные данные."
    )
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": f"Анализ патогена: {name}"}],
        model=GROQ_MODEL_TEXT, temperature=0.1,
    )


async def medical_database_query(query: str) -> str:
    """Запрос к медицинской базе знаний"""
    system = (
        "Ты медицинский информационный ассистент с доступом к базе медицинских знаний. "
        "Давай информацию о: болезнях, симптомах, лекарствах, процедурах, анализах.\n"
        "ВАЖНО: Всегда добавляй disclaimer о необходимости консультации врача.\n"
        "Используй: МКБ-10 коды, фармакологические термины, клинические протоколы."
    )
    result = await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": query}],
        model=GROQ_MODEL_TEXT, temperature=0.1,
    )
    return result + "\n\n⚠️ Это справочная информация. Обязательно консультируйтесь с врачом."


async def check_health_dangers(metrics: Dict[str, float]) -> List[str]:
    """Проверка показателей здоровья на опасные значения"""
    warnings = []
    if "pulse" in metrics:
        p = metrics["pulse"]
        if p > 100:
            warnings.append(f"⚠️ ПУЛЬС {p} уд/мин — тахикардия. Обратитесь к врачу.")
        elif p < 50:
            warnings.append(f"⚠️ ПУЛЬС {p} уд/мин — брадикардия. Обратитесь к врачу.")
    if "pressure_sys" in metrics:
        ps = metrics["pressure_sys"]
        if ps > 140:
            warnings.append(f"🚨 ДАВЛЕНИЕ {ps} мм рт.ст. — гипертония. Срочно к врачу!")
        elif ps < 90:
            warnings.append(f"⚠️ ДАВЛЕНИЕ {ps} мм рт.ст. — гипотония.")
    if "temperature" in metrics:
        t = metrics["temperature"]
        if t >= 38.5:
            warnings.append(f"🚨 ТЕМПЕРАТУРА {t}°C — высокая лихорадка. Обратитесь к врачу!")
        elif t >= 37.5:
            warnings.append(f"⚠️ ТЕМПЕРАТУРА {t}°C — субфебрильная.")
    if "sleep" in metrics:
        s = metrics["sleep"]
        if s < 5:
            warnings.append(f"⚠️ СОН {s} ч — критически мало. Риск для здоровья.")
    if "glucose" in metrics:
        g = metrics["glucose"]
        if g > 7.8:
            warnings.append(f"⚠️ ГЛЮКОЗА {g} ммоль/л — выше нормы.")
        elif g < 3.5:
            warnings.append(f"🚨 ГЛЮКОЗА {g} ммоль/л — гипогликемия!")
    return warnings


# =========================
# NAVIGATION
# =========================

async def build_route(from_place: str, to_place: str, mode: str = "driving") -> str:
    """Построение маршрута"""
    system = (
        "Ты навигационный ассистент. Дай подробный маршрут:\n"
        "• Оптимальный маршрут с пошаговыми инструкциями\n"
        "• Примерное время в пути\n"
        "• Альтернативные маршруты\n"
        "• Возможные проблемные участки\n"
        "• Безопасный маршрут (объезд опасных зон)"
    )
    mode_ru = {"driving": "на автомобиле", "walking": "пешком", "transit": "на общественном транспорте"}.get(mode, mode)
    return await groq_chat(
        [
            {"role": "system", "content": system},
            {"role": "user", "content": f"Маршрут {mode_ru} из {from_place} в {to_place}"},
        ],
        model=GROQ_MODEL_FAST, temperature=0.3,
    )


def parse_coordinates(text: str) -> Optional[Dict[str, float]]:
    """Определение координат из текста"""
    patterns = [
        r"(\d{1,3}[.,]\d+)[°\s]*[NСсн]\w*[,\s]+(\d{1,3}[.,]\d+)[°\s]*[EВвe]",
        r"(\d{1,3}[.,]\d+)\s*,\s*(\d{1,3}[.,]\d+)",
        r"широта\s*:?\s*(\d{1,3}[.,]\d+)[,\s]+долгота\s*:?\s*(\d{1,3}[.,]\d+)",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            lat = float(m.group(1).replace(",", "."))
            lon = float(m.group(2).replace(",", "."))
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                return {"lat": lat, "lon": lon}
    return None


# =========================
# BEHAVIORAL ANALYSIS
# =========================

async def analyze_behavior(text: str, user_id: int) -> str:
    """Анализ поведения человека из текста"""
    system = (
        "Ты психолог и поведенческий аналитик. Проанализируй текст/описание поведения.\n"
        "Определи:\n"
        "• Доминирующие поведенческие паттерны\n"
        "• Эмоциональное состояние\n"
        "• Мотивация и скрытые цели\n"
        "• Уровень стресса (1-10)\n"
        "• Тип личности (по MBTI или другой модели)\n"
        "• Прогноз поведения\n"
        "• Рекомендации по взаимодействию\n"
        "Будь объективным, не давай оценочных суждений."
    )
    result = await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": text}],
        model=GROQ_MODEL_TEXT, temperature=0.3,
    )

    # Сохраняем паттерны
    patterns_found = re.findall(r"паттерн[:\s]+([^\n]+)", result, re.I)
    for p in patterns_found[:3]:
        await update_behavior_pattern(user_id, p.strip()[:100])

    return result


async def analyze_social_connections(description: str) -> str:
    """Анализ социальных связей"""
    system = (
        "Ты аналитик социальных сетей и связей. Проанализируй описание.\n"
        "Определи:\n"
        "• Ключевые узлы влияния\n"
        "• Скрытые зависимости\n"
        "• Потенциальные конфликты\n"
        "• Сильные и слабые связи\n"
        "• Рычаги влияния\n"
        "• Риски"
    )
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": description}],
        model=GROQ_MODEL_TEXT, temperature=0.3,
    )


# =========================
# ANALYTICS & FORECASTING
# =========================

def linear_regression(x_data: List[float], y_data: List[float]) -> Dict[str, Any]:
    """Линейная регрессия"""
    if len(x_data) != len(y_data) or len(x_data) < 2:
        return {"error": "Нужно минимум 2 точки, x и y одинаковой длины"}
    n = len(x_data)
    sum_x = sum(x_data)
    sum_y = sum(y_data)
    sum_xy = sum(x_data[i] * y_data[i] for i in range(n))
    sum_xx = sum(x ** 2 for x in x_data)
    denom = n * sum_xx - sum_x ** 2
    if denom == 0:
        return {"error": "Деление на ноль при регрессии"}
    k = (n * sum_xy - sum_x * sum_y) / denom
    b = (sum_y - k * sum_x) / n
    y_pred = [k * x + b for x in x_data]
    ss_res = sum((y_data[i] - y_pred[i]) ** 2 for i in range(n))
    ss_tot = sum((y - sum_y / n) ** 2 for y in y_data)
    r2 = 1 - ss_res / ss_tot if ss_tot != 0 else 1.0
    return {
        "наклон (k)": round(k, 4),
        "сдвиг (b)": round(b, 4),
        "формула": f"y = {round(k,4)}x + {round(b,4)}",
        "R²": round(r2, 4),
        "качество": "отлично" if r2 > 0.9 else "хорошо" if r2 > 0.7 else "средне" if r2 > 0.5 else "слабо",
    }


def forecast_next(values: List[float], steps: int = 3) -> Dict[str, Any]:
    """Прогнозирование следующих значений временного ряда"""
    if len(values) < 3:
        return {"error": "Нужно минимум 3 значения"}
    x_data = list(range(len(values)))
    reg = linear_regression(x_data, values)
    if "error" in reg:
        return reg
    k = reg["наклон (k)"]
    b = reg["сдвиг (b)"]
    forecasts = []
    for i in range(1, steps + 1):
        x_next = len(values) - 1 + i
        forecasts.append(round(k * x_next + b, 4))
    trend = "растёт" if k > 0 else "падает" if k < 0 else "стабильно"
    return {
        "прогноз": forecasts,
        "тренд": trend,
        "наклон": k,
        "R²": reg["R²"],
    }


async def analyze_risk(scenario: str) -> str:
    """Оценка рисков"""
    system = (
        "Ты риск-аналитик. Оцени сценарий по методологии:\n"
        "• Идентификация рисков\n"
        "• Вероятность реализации (1-10)\n"
        "• Возможный ущерб (1-10)\n"
        "• Итоговая оценка риска = вероятность × ущерб\n"
        "• Меры по снижению рисков\n"
        "• Резервные планы\n"
        "Дай числовые оценки и конкретные рекомендации."
    )
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": scenario}],
        model=GROQ_MODEL_TEXT, temperature=0.2,
    )


async def generate_analytical_report(user_id: int) -> str:
    """Генерация аналитического отчёта"""
    events = await db_query(
        "SELECT event_type, COUNT(*) as cnt FROM analytics_events WHERE user_id = ? AND created_at >= datetime('now', '-30 day') GROUP BY event_type ORDER BY cnt DESC",
        (user_id,),
    )
    finance = await get_finance_records(user_id, limit=100)
    health = await db_query(
        "SELECT metric, AVG(value) as avg_val, COUNT(*) as cnt FROM health_metrics WHERE user_id = ? AND created_at >= datetime('now', '-30 day') GROUP BY metric",
        (user_id,),
    )
    threats = await db_query(
        "SELECT threat_type, severity, created_at FROM threat_log WHERE user_id = ? ORDER BY created_at DESC LIMIT 5",
        (user_id,),
    )
    behaviors = await get_behavior_patterns(user_id)

    lines = [f"📊 АНАЛИТИЧЕСКИЙ ОТЧЁТ — {now_str()}", ""]

    if events:
        lines.append("📈 Активность (30 дней):")
        for e in events[:5]:
            lines.append(f"  • {e['event_type']}: {e['cnt']} раз")
        lines.append("")

    if health:
        lines.append("💊 Здоровье (30 дней):")
        for h in health:
            if h["avg_val"] is not None:
                lines.append(f"  • {h['metric']}: среднее {round(h['avg_val'], 1)} ({h['cnt']} записей)")
        lines.append("")

    if finance:
        income = sum(f["amount"] for f in finance if f["record_type"] == "income")
        expense = sum(f["amount"] for f in finance if f["record_type"] == "expense")
        lines.append(f"💰 Финансы: доход {income:.0f} | расходы {expense:.0f} | баланс {income - expense:.0f}")
        lines.append("")

    if threats:
        lines.append("🛡 Угрозы:")
        for t in threats:
            lines.append(f"  • [{t['severity']}/10] {t['threat_type']}: {t['description'][:80]}")
        lines.append("")

    if behaviors:
        lines.append("🧠 Поведенческие паттерны:")
        for b in behaviors[:5]:
            lines.append(f"  • {b['pattern']} (×{b['count']})")

    return "\n".join(lines) if len(lines) > 3 else "Недостаточно данных для отчёта."


# =========================
# FINANCE
# =========================

async def finance_analysis(user_id: int) -> str:
    records = await get_finance_records(user_id, limit=200)
    if not records:
        return "Финансовых данных нет. Добавь записи: /income 5000 зарплата или /expense 200 продукты"

    income_by_cat: Dict[str, float] = {}
    expense_by_cat: Dict[str, float] = {}
    income_total = 0.0
    expense_total = 0.0

    for r in records:
        cat = r.get("category") or "прочее"
        amount = float(r.get("amount", 0))
        if r["record_type"] == "income":
            income_total += amount
            income_by_cat[cat] = income_by_cat.get(cat, 0) + amount
        elif r["record_type"] == "expense":
            expense_total += amount
            expense_by_cat[cat] = expense_by_cat.get(cat, 0) + amount

    balance = income_total - expense_total
    savings_rate = (balance / income_total * 100) if income_total > 0 else 0

    lines = [
        f"💰 ФИНАНСОВЫЙ АНАЛИЗ",
        f"Доходы: {income_total:,.0f}",
        f"Расходы: {expense_total:,.0f}",
        f"Баланс: {balance:,.0f} ({'✅' if balance >= 0 else '🔴'})",
        f"Норма сбережений: {savings_rate:.1f}%",
        "",
    ]

    if income_by_cat:
        lines.append("📈 Доходы по категориям:")
        for cat, amount in sorted(income_by_cat.items(), key=lambda x: x[1], reverse=True)[:5]:
            lines.append(f"  • {cat}: {amount:,.0f}")

    if expense_by_cat:
        lines.append("\n📉 Расходы по категориям:")
        for cat, amount in sorted(expense_by_cat.items(), key=lambda x: x[1], reverse=True)[:5]:
            pct = amount / expense_total * 100 if expense_total > 0 else 0
            lines.append(f"  • {cat}: {amount:,.0f} ({pct:.1f}%)")

    if savings_rate < 10:
        lines.append("\n⚠️ Норма сбережений ниже 10% — рекомендую оптимизировать расходы.")
    elif savings_rate > 30:
        lines.append("\n✅ Отличная норма сбережений!")

    return "\n".join(lines)


async def forecast_budget(user_id: int, months_ahead: int = 3) -> str:
    records = await get_finance_records(user_id, limit=200)
    if len(records) < 5:
        return "Недостаточно данных для прогноза. Нужно минимум 5 записей."

    monthly: Dict[str, Dict[str, float]] = {}
    for r in records:
        month = r["created_at"][:7]
        if month not in monthly:
            monthly[month] = {"income": 0.0, "expense": 0.0}
        if r["record_type"] == "income":
            monthly[month]["income"] += float(r["amount"])
        elif r["record_type"] == "expense":
            monthly[month]["expense"] += float(r["amount"])

    months_sorted = sorted(monthly.keys())
    if len(months_sorted) < 2:
        return "Нужны данные минимум за 2 месяца."

    incomes = [monthly[m]["income"] for m in months_sorted]
    expenses = [monthly[m]["expense"] for m in months_sorted]

    inc_forecast = forecast_next(incomes, months_ahead)
    exp_forecast = forecast_next(expenses, months_ahead)

    lines = [f"📅 ПРОГНОЗ БЮДЖЕТА на {months_ahead} месяца:"]
    if "error" not in inc_forecast and "error" not in exp_forecast:
        for i in range(months_ahead):
            inc = inc_forecast["прогноз"][i]
            exp = exp_forecast["прогноз"][i]
            bal = inc - exp
            lines.append(f"  Месяц +{i+1}: доход ~{inc:,.0f} | расходы ~{exp:,.0f} | баланс ~{bal:,.0f}")
        lines.append(f"\nТренд доходов: {inc_forecast['тренд']}")
        lines.append(f"Тренд расходов: {exp_forecast['тренд']}")
    else:
        lines.append("Не удалось построить прогноз.")

    return "\n".join(lines)


# =========================
# SELF-REPAIR
# =========================

async def self_repair_check() -> str:
    """Самодиагностика и самовосстановление"""
    issues = []
    fixes = []

    # Проверка DB
    try:
        await db_query("SELECT COUNT(*) FROM settings")
    except Exception as e:
        issues.append(f"DB недоступна: {e}")
        try:
            init_db()
            fixes.append("DB пересоздана")
            log_self_repair("self_repair", f"DB error: {e}", "reinit_db")
        except Exception as e2:
            fixes.append(f"DB восстановить не удалось: {e2}")

    # Проверка HTTP session
    try:
        session = await get_http_session()
        if session.closed:
            issues.append("HTTP session закрыта")
            global HTTP_SESSION
            HTTP_SESSION = None
            await get_http_session()
            fixes.append("HTTP session пересоздана")
            log_self_repair("self_repair", "HTTP session closed", "recreated")
    except Exception as e:
        issues.append(f"HTTP session error: {e}")

    # Проверка Groq
    try:
        test_resp = await groq_chat(
            [{"role": "user", "content": "ping"}],
            model=GROQ_MODEL_FAST, temperature=0.0,
        )
        if "Ошибка" in test_resp or "недоступен" in test_resp:
            issues.append("Groq API не отвечает корректно")
        else:
            fixes.append("Groq API — ОК")
    except Exception as e:
        issues.append(f"Groq ping failed: {e}")

    # Проверка DATA_DIR
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR, exist_ok=True)
        fixes.append("DATA_DIR создан")
        log_self_repair("self_repair", "DATA_DIR missing", "created")

    recent_repairs = SELF_REPAIR_LOG[-5:] if SELF_REPAIR_LOG else []

    lines = ["🔧 САМОДИАГНОСТИКА JARVIS"]
    lines.append(f"Время: {now_str()}")
    lines.append(f"Проблем найдено: {len(issues)}")
    lines.append(f"Исправлений применено: {len(fixes)}")
    if issues:
        lines.append("\n⚠️ Проблемы:")
        for i in issues:
            lines.append(f"  • {i}")
    if fixes:
        lines.append("\n✅ Действия:")
        for f in fixes:
            lines.append(f"  • {f}")
    if recent_repairs:
        lines.append("\n📋 Последние авторемонты:")
        for r in recent_repairs:
            lines.append(f"  [{r['time']}] {r['module']}: {r['fix']}")

    status = "✅ Всё в порядке" if not issues else f"⚠️ Найдено проблем: {len(issues)}, исправлено: {len(fixes)}"
    lines.append(f"\nСтатус: {status}")
    return "\n".join(lines)


# =========================
# TRANSLATE
# =========================

async def translate_text(text: str, target_lang: str) -> str:
    system = "Ты профессиональный переводчик. Переводи на указанный язык. Верни только готовый перевод, без пояснений."
    return await groq_chat(
        [
            {"role": "system", "content": system},
            {"role": "user", "content": f"Язык: {target_lang}\n\nТекст:\n{text}"},
        ],
        model=GROQ_MODEL_FAST, temperature=0.2,
    )


async def detect_language(text: str) -> str:
    result = await groq_chat(
        [
            {"role": "system", "content": "Определи язык текста. Верни только название языка на русском, одно слово."},
            {"role": "user", "content": text[:500]},
        ],
        model=GROQ_MODEL_FAST, temperature=0.0,
    )
    return result.strip()


# =========================
# LLM LEARNING
# =========================

async def teach_topic(topic: str, level: str = "средний") -> str:
    """Обучение по теме"""
    system = (
        f"Ты опытный преподаватель. Уровень ученика: {level}.\n"
        "Объясни тему:\n"
        "• Простое введение\n"
        "• Ключевые концепции\n"
        "• Примеры из жизни\n"
        "• Практические задания\n"
        "• Что изучить дальше\n"
        "Адаптируй язык к уровню. Используй аналогии."
    )
    return await groq_chat(
        [{"role": "system", "content": system}, {"role": "user", "content": f"Объясни: {topic}"}],
        model=GROQ_MODEL_TEXT, temperature=0.4,
    )


async def llm_extract_facts(text: str) -> Dict[str, Any]:
    system = (
        "Ты извлекаешь устойчивые факты из сообщений пользователя для долгосрочной памяти. "
        "Верни СТРОГО JSON без markdown. "
        "Формат: {\"facts\": [\"...\"], \"health\": [{\"metric\":\"sleep|smoke|mood\", \"value\": 7, \"text_value\":\"...\", \"note\":\"...\", \"meta\": {}}], \"followup_suggestion\": \"...\"}\n"
        "Сохраняй только полезные, устойчивые факты. Не сохраняй мусор."
    )
    raw = await groq_chat(
        [
            {"role": "system", "content": system},
            {"role": "user", "content": clamp_text(text, 6000)},
        ],
        model=GROQ_MODEL_FAST, temperature=0.0,
    )
    return safe_json_loads(raw, {"facts": [], "health": [], "followup_suggestion": ""})


def detect_health_from_text(text: str) -> List[Dict[str, Any]]:
    t = text.lower()
    result: List[Dict[str, Any]] = []
    m = re.search(r"(?:спал|сон)[^\d]{0,20}(\d{1,2}:\d{2})\s*(?:до|-)\s*(\d{1,2}:\d{2})", t, re.I)
    if m:
        hours = calc_sleep_hours(m.group(1), m.group(2))
        result.append({"metric": "sleep", "value": hours, "text_value": f"{m.group(1)}-{m.group(2)}", "note": "сон по сообщению", "meta": {}})
    m = re.search(r"(?:спал|сон)[^\d]{0,20}(\d+(?:[.,]\d+)?)\s*(?:ч|час|часа|часов)", t, re.I)
    if m:
        result.append({"metric": "sleep", "value": float(m.group(1).replace(",", ".")), "text_value": m.group(1), "note": "длительность сна", "meta": {}})
    smoke_count = 0
    m = re.search(r"(?:выкурил(?:а)?|сигарет(?:а|ы)?|курил(?:а)?)[^\d]{0,20}\+?(\d+)", t, re.I)
    if m:
        smoke_count = int(m.group(1))
    elif re.search(r"\bкурил\b|\bсигарет\b", t, re.I):
        smoke_count = 1
    if smoke_count > 0:
        result.append({"metric": "smoke", "value": float(smoke_count), "text_value": str(smoke_count), "note": "курение", "meta": {}})
    m = re.search(r"(?:настроение|самочувствие)\s*[:=]?\s*(\d{1,2})(?:/10)?", t, re.I)
    if m:
        result.append({"metric": "mood", "value": float(m.group(1)), "text_value": m.group(1), "note": "самочувствие", "meta": {}})

    # Пульс
    m = re.search(r"(?:пульс|чсс)[^\d]{0,10}(\d{2,3})", t, re.I)
    if m:
        result.append({"metric": "pulse", "value": float(m.group(1)), "text_value": m.group(1), "note": "пульс", "meta": {}})

    # Давление
    m = re.search(r"давление[^\d]{0,10}(\d{2,3})\s*/\s*(\d{2,3})", t, re.I)
    if m:
        result.append({"metric": "pressure_sys", "value": float(m.group(1)), "text_value": f"{m.group(1)}/{m.group(2)}", "note": "давление", "meta": {"dia": m.group(2)}})

    # Температура
    m = re.search(r"температура[^\d]{0,10}(\d{2}[.,]\d)", t, re.I)
    if m:
        result.append({"metric": "temperature", "value": float(m.group(1).replace(",", ".")), "text_value": m.group(1), "note": "температура", "meta": {}})

    return result


async def auto_learn_from_text(user_id: int, text: str, source: str = "chat") -> None:
    if len(text.strip()) < 8:
        return
    health_items = detect_health_from_text(text)
    for item in health_items:
        await add_health_metric(user_id, metric=item["metric"], value=item.get("value"), text_value=item.get("text_value"), note=item.get("note"), meta=item.get("meta"))

    # Предупреждения об опасности
    if health_items:
        metrics_dict = {item["metric"]: item["value"] for item in health_items if item.get("value")}
        warnings = await check_health_dangers(metrics_dict)
        if warnings:
            for w in warnings:
                await log_threat(user_id, "health_danger", w, severity=7)

    try:
        extracted = await llm_extract_facts(text)
    except Exception:
        return

    for fact in (extracted.get("facts") or [])[:10]:
        fact = normalize_text(str(fact))
        if len(fact) >= 4:
            await add_memory(user_id, fact)

    for h in (extracted.get("health") or [])[:10]:
        metric = str(h.get("metric") or "").strip().lower()
        if metric in {"sleep", "smoke", "mood", "pulse", "pressure_sys", "temperature"}:
            await add_health_metric(user_id, metric=metric, value=h.get("value"), text_value=h.get("text_value"), note=h.get("note"), meta=h.get("meta") or {})

    await log_analytics_event(user_id, "text_processed", {"source": source, "length": len(text)})


async def build_health_summary(user_id: int) -> str:
    rows = await db_query(
        "SELECT metric, value, created_at FROM health_metrics WHERE user_id = ? AND created_at >= datetime('now', '-14 day') ORDER BY id DESC",
        (user_id,),
    )
    if not rows:
        return "Health data: нет данных."
    sleep_vals = [float(r["value"]) for r in rows if r["metric"] == "sleep" and r["value"] is not None]
    mood_vals = [float(r["value"]) for r in rows if r["metric"] == "mood" and r["value"] is not None]
    smoke_total = sum(float(r["value"]) for r in rows if r["metric"] == "smoke" and r["value"] is not None)
    pulse_vals = [float(r["value"]) for r in rows if r["metric"] == "pulse" and r["value"] is not None]
    parts = []
    if sleep_vals:
        parts.append(f"Сон: {round(sum(sleep_vals)/len(sleep_vals), 1)} ч")
    if mood_vals:
        parts.append(f"Настроение: {round(sum(mood_vals)/len(mood_vals), 1)}/10")
    if smoke_total:
        parts.append(f"Курение: {int(smoke_total)}")
    if pulse_vals:
        parts.append(f"Пульс: {round(sum(pulse_vals)/len(pulse_vals), 0)}")
    return " | ".join(parts) if parts else "Health data: недостаточно данных."


async def build_proactive_advice(user_id: int) -> Optional[str]:
    last_msgs = await db_query(
        "SELECT content FROM messages WHERE user_id = ? AND role = 'user' ORDER BY id DESC LIMIT 30",
        (user_id,),
    )
    joined = " ".join(r["content"].lower() for r in last_msgs)
    fatigue_hits = sum(1 for k in ["устал", "не высп", "нет сил", "сонлив"] if k in joined)
    if fatigue_hits >= 2:
        return "😴 Заметил признаки усталости. Хочешь план восстановления?"

    health_rows = await db_query(
        "SELECT metric, value FROM health_metrics WHERE user_id = ? AND created_at >= datetime('now', '-7 day')",
        (user_id,),
    )
    sleep_vals = [float(r["value"]) for r in health_rows if r["metric"] == "sleep" and r["value"] is not None]
    if sleep_vals and (sum(sleep_vals) / len(sleep_vals)) < 6.5:
        return "💤 Сон в среднем меньше нормы. Хочешь рекомендации?"

    threats = await db_query(
        "SELECT COUNT(*) as cnt FROM threat_log WHERE user_id = ? AND created_at >= datetime('now', '-1 day') AND resolved = 0",
        (user_id,),
    )
    if threats and threats[0]["cnt"] > 0:
        return f"🛡 Есть {threats[0]['cnt']} необработанных угроз за сутки. /threats"

    return None


# =========================
# PROMPTS
# =========================

async def build_system_prompt(user_id: int, query: str) -> str:
    settings = await get_settings(user_id)
    relevant = await search_knowledge(query, user_id, top_k=3, scope="all")
    context_block = "\n".join(
        [f"[{i['source']} #{i['id']} | {i['score']:.3f}] {i['text']}" for i in relevant]
    ) or "Нет релевантного контекста."
    tasks = await get_tasks(user_id, status="open", limit=5)
    task_block = "\n".join([f"- #{t['id']} [P{t.get('priority',2)}]: {t['text']}" for t in tasks]) or "Нет открытых задач."
    health_block = await build_health_summary(user_id)
    behaviors = await get_behavior_patterns(user_id)
    behavior_block = ", ".join([b["pattern"] for b in behaviors[:5]]) or "нет данных"
    city = settings.get("default_city") or "не задан"

    return (
        "Ты Jarvis — продвинутый личный ИИ-ассистент.\n"
        "Возможности: анализ угроз, поведенческий анализ, научные расчёты, медицина, финансы, тактика, обучение.\n"
        "Стиль: умный, прямой, иногда с сарказмом, без воды. Используй эмодзи умеренно.\n"
        "Можешь шутить и быть саркастичным, когда уместно.\n"
        "Делай самостоятельные выводы на основе данных пользователя.\n"
        "Если видишь паттерн или риск — предупреди без лишних слов.\n\n"
        f"Время: {now_str()} | Город: {city}\n"
        f"Модель: {settings.get('model_mode', 'smart')} | Voice: {'ON' if int(settings.get('voice_mode', 0)) else 'OFF'}\n\n"
        f"Задачи:\n{task_block}\n\n"
        f"Здоровье: {health_block}\n\n"
        f"Поведенческие паттерны: {behavior_block}\n\n"
        f"Контекст:\n{context_block}"
    )


async def build_briefing(user_id: int) -> str:
    settings = await get_settings(user_id)
    city = settings.get("default_city")
    lines = [f"☀️ Брифинг — {now_str()}"]
    if city:
        lines.append(await get_weather_text(city))
    else:
        lines.append("Погода: город не задан.")
    open_tasks = await get_tasks(user_id, status="open", limit=7)
    if open_tasks:
        lines.append("\n📌 Задачи:")
        for t in open_tasks:
            due = f" — до {t['due_at']}" if t.get("due_at") else ""
            priority_icon = "🔴" if t.get("priority") == 1 else "🟡" if t.get("priority") == 2 else "🟢"
            lines.append(f"  {priority_icon} #{t['id']} {t['text']}{due}")
    else:
        lines.append("\n📌 Задачи: нет открытых.")
    due_reminders = await get_due_reminders_for_user(user_id)
    if due_reminders:
        lines.append("\n⏰ Напоминания:")
        for r in due_reminders[:5]:
            lines.append(f"  • {r['text']} ({r['due_at']})")
    threats = await db_query(
        "SELECT COUNT(*) as cnt FROM threat_log WHERE user_id = ? AND resolved = 0 AND created_at >= datetime('now', '-1 day')",
        (user_id,),
    )
    if threats and threats[0]["cnt"] > 0:
        lines.append(f"\n🛡 Активных угроз: {threats[0]['cnt']}")
    health = await build_health_summary(user_id)
    lines.append(f"\n💊 Здоровье: {health}")
    finance = await finance_analysis(user_id)
    if "Финансовых данных нет" not in finance:
        first_line = finance.split("\n")[0]
        lines.append(f"\n{first_line}")
    return "\n".join(lines)


# =========================
# PARSERS
# =========================

def parse_relative_due(expr: str) -> Optional[datetime]:
    m = re.match(r"^(\d+)\s*([smhd])$", expr.strip().lower())
    if not m:
        return None
    value = int(m.group(1))
    unit = m.group(2)
    delta = {"s": timedelta(seconds=value), "m": timedelta(minutes=value), "h": timedelta(hours=value), "d": timedelta(days=value)}.get(unit)
    return datetime.now() + delta if delta else None


def parse_absolute_due(expr: str) -> Optional[datetime]:
    expr = expr.strip()
    for fmt in ("%Y-%m-%d %H:%M", "%d.%m.%Y %H:%M", "%d.%m %H:%M", "%d.%m.%Y", "%d.%m"):
        try:
            dt = datetime.strptime(expr, fmt)
            if fmt in {"%d.%m", "%d.%m.%Y"}:
                dt = dt.replace(year=datetime.now().year, hour=9, minute=0)
            return dt
        except Exception:
            pass
    m = re.match(r"^(\d{1,2})\s+([а-яё]+)(?:\s+(\d{4}))?(?:\s+в\s+(\d{1,2}:\d{2}))?$", expr.lower())
    if m:
        months = {"января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6, "июля": 7, "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12}
        month_name = m.group(2)
        if month_name not in months:
            return None
        hhmm = m.group(4) or "09:00"
        hh, mm = map(int, hhmm.split(":"))
        try:
            return datetime(int(m.group(3) or datetime.now().year), months[month_name], int(m.group(1)), hh, mm)
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
            return (datetime.now() + timedelta(days=1)).replace(hour=hh, minute=mm, second=0, microsecond=0)
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
    due = parse_due_from_text(m.group(2).strip())
    if not due or len(task_text) < 3:
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
    return None


def detect_math_request(text: str) -> Optional[str]:
    math_patterns = [
        r"вычисли|посчитай|рассчитай|сколько будет",
        r"[\d+\-*/()^√]+",
        r"формула|уравнение|интеграл|производная",
    ]
    t = text.lower().strip()
    if any(re.search(p, t) for p in math_patterns):
        expr_match = re.search(r"[\d\s+\-*/().^√sincotaglqrb]+", text)
        if expr_match:
            return expr_match.group(0).strip()
    return None


# =========================
# CONFIRMED ACTIONS
# =========================

async def execute_confirmed_action(action_type: str, payload: Dict[str, Any], user_id: int) -> str:
    if action_type == "delete_memory":
        await delete_memory(user_id, int(payload["memory_id"]))
        return f"Память #{payload['memory_id']} удалена."
    if action_type == "reset_history":
        await clear_history(user_id)
        return "История очищена."
    if action_type == "set_city":
        await set_setting(user_id, "default_city", payload["city"])
        return f"Город: {payload['city']}"
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
    if action_type == "set_setting":
        await set_setting(user_id, payload["key"], payload["value"])
        return "Настройка изменена."
    return "Готово."


async def ask_confirmation(message: Message, title: str, action_type: str, payload: Dict[str, Any]) -> None:
    cid = await add_confirmation(message.from_user.id, action_type, payload)
    kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Да", callback_data=f"confirm:{cid}"),
        InlineKeyboardButton(text="❌ Нет", callback_data=f"cancel:{cid}"),
    ]])
    await message.answer(f"{title}\n\nПодтвердить?", reply_markup=kb)


@router.callback_query(F.data.startswith("confirm:"))
async def cb_confirm_action(callback: CallbackQuery) -> None:
    user_id = callback.from_user.id
    cid = int(callback.data.split(":", 1)[1])
    row = await get_confirmation(cid)
    if not row or int(row["user_id"]) != user_id:
        await callback.answer("Не найдено или устарело.", show_alert=True)
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
    await callback.message.answer("Отменено.")


# =========================
# INTENT DETECTION
# =========================

async def handle_natural_language_intents(message: Message, text: str) -> Optional[str]:
    user_id = message.from_user.id
    t = normalize_text(text)

    # Перевод
    tr = detect_translation_request(t)
    if tr:
        target_lang, source_text = tr
        if not source_text:
            return "Напиши текст для перевода."
        return await translate_text(source_text, target_lang)

    # Математика
    if any(kw in t.lower() for kw in ["вычисли", "посчитай", "рассчитай", "сколько будет"]):
        expr = re.sub(r"(вычисли|посчитай|рассчитай|сколько будет)\s*", "", t, flags=re.I).strip()
        if expr:
            return calculate_math(expr)

    # Веб-поиск
    if re.match(r"^(найди|поищи|search|загугли)\s+", t, re.I):
        query = re.sub(r"^(найди|поищи|search|загугли)\s+", "", t, flags=re.I).strip()
        return await web_search(query)

    # Новости
    if re.match(r"^(новости|news)\s*", t, re.I):
        query = re.sub(r"^(новости|news)\s*", "", t, flags=re.I).strip()
        return await fetch_news(query)

    # Анализ угроз
    if re.match(r"^(анализ угрозы|угроза|опасность|threat)\s+", t, re.I):
        desc = re.sub(r"^(анализ угрозы|угроза|опасность|threat)\s+", "", t, flags=re.I).strip()
        return await analyze_threat(desc, user_id)

    # Тактика
    if re.match(r"^(тактика|тактические|стратегия)\s+", t, re.I):
        return await tactical_recommendations(t)

    # Поведенческий анализ
    if re.match(r"^(анализ поведения|психология|поведение человека)\s+", t, re.I):
        desc = re.sub(r"^(анализ поведения|психология|поведение человека)\s+", "", t, flags=re.I).strip()
        return await analyze_behavior(desc, user_id)

    # Химия
    if re.match(r"^(химия|химический анализ|состав)\s+", t, re.I):
        substance = re.sub(r"^(химия|химический анализ|состав)\s+", "", t, flags=re.I).strip()
        return await chemical_model(substance)

    # Физика
    if re.match(r"^(физика|физическая модель)\s+", t, re.I):
        return await physical_model(re.sub(r"^(физика|физическая модель)\s+", "", t, flags=re.I).strip())

    # Вирус / медицина
    if re.match(r"^(вирус|патоген|анализ вируса)\s+", t, re.I):
        return await analyze_virus(re.sub(r"^(вирус|патоген|анализ вируса)\s+", "", t, flags=re.I).strip())

    if re.match(r"^(медицина|болезнь|симптомы|лекарство)\s+", t, re.I):
        return await medical_database_query(t)

    # Маршрут
    if re.match(r"^(маршрут|навигация|как добраться)\s+", t, re.I):
        parts_nav = re.split(r"\s+до\s+|\s*->\s*|\s+в\s+", t, maxsplit=1)
        if len(parts_nav) >= 2:
            from_place = re.sub(r"^(маршрут|навигация|как добраться)\s+(из\s+|от\s+)?", "", parts_nav[0], flags=re.I).strip()
            to_place = parts_nav[1].strip()
            return await build_route(from_place, to_place)
        return "Укажи маршрут: маршрут из [откуда] до [куда]"

    # Финансовый анализ
    if re.search(r"\b(финансы|бюджет|расходы|доходы|финансовый анализ)\b", t, re.I):
        return await finance_analysis(user_id)

    # Обучение
    if re.match(r"^(объясни|расскажи про|научи|обучи)\s+", t, re.I):
        topic = re.sub(r"^(объясни|расскажи про|научи|обучи)\s+", "", t, flags=re.I).strip()
        return await teach_topic(topic)

    # Напоминания
    reminder = parse_natural_reminder(t)
    if reminder:
        due = reminder["due"].isoformat(timespec="minutes")
        if reminder.get("mode") == "recurring":
            rid = await add_recurring_reminder(user_id=user_id, chat_id=message.chat.id, text=reminder["text"], rule=reminder["rule"], next_due=reminder["due"])
            return f"Периодическое напоминание #{rid}\nПервый запуск: {due}"
        rid = await add_reminder(user_id, message.chat.id, reminder["text"], due)
        return f"⏰ Напоминание #{rid}\nВремя: {due}"

    # Задача с дедлайном
    task_with_due = parse_task_with_deadline(t)
    if task_with_due:
        priority = 2
        for kw, p in PRIORITY_MAP.items():
            if kw in t.lower():
                priority = p
                break
        task_id = await add_task(user_id, task_with_due["task_text"], task_with_due["due"].isoformat(timespec="minutes"), priority)
        due = task_with_due["due"]
        remind_at = due - timedelta(hours=1)
        if remind_at > datetime.now():
            await add_reminder(user_id, message.chat.id, f"Задача #{task_id}: {task_with_due['task_text']}", remind_at.isoformat(timespec="minutes"))
        return f"✅ Задача #{task_id}\nДедлайн: {format_dt(due)}"

    # Удаление/очистка
    if re.search(r"\b(удали|сотри|очисти)\b", t, re.I):
        if re.search(r"\bисторию\b|\bпамять\b", t, re.I):
            await ask_confirmation(message, "Очистить историю сообщений?", "reset_history", {})
            return ""
        m = re.search(r"(?:удали|сотри)\s+(?:память|memory)\s*(\d+)", t, re.I)
        if m:
            await ask_confirmation(message, f"Удалить память #{m.group(1)}?", "delete_memory", {"memory_id": int(m.group(1))})
            return ""

    # Настройки
    if re.search(r"\b(выключи|включи|поменяй|измени|установи)\b", t, re.I):
        if re.search(r"\bголос\b", t, re.I):
            new_value = 0 if re.search(r"\bвыключи\b", t, re.I) else 1
            await ask_confirmation(message, f"Voice mode → {'ON' if new_value else 'OFF'}?", "set_voice", {"value": new_value})
            return ""
        if re.search(r"\bбрифинг\b", t, re.I):
            new_value = 0 if re.search(r"\bвыключи\b", t, re.I) else 1
            await ask_confirmation(message, f"Briefing → {'ON' if new_value else 'OFF'}?", "set_briefing", {"value": new_value})
            return ""
        if re.search(r"\bмодель\b", t, re.I):
            m2 = re.search(r"\b(fast|smart|vision)\b", t, re.I)
            if m2:
                await ask_confirmation(message, f"Model → {m2.group(1).lower()}?", "set_model", {"value": m2.group(1).lower()})
                return ""
        if re.search(r"\bгород\b", t, re.I):
            m3 = re.search(r"(?:на|в)\s+([А-Яа-яЁёA-Za-z\- ]+)$", t)
            if m3:
                city = normalize_text(m3.group(1))
                await ask_confirmation(message, f"Город → «{city}»?", "set_city", {"city": city})
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
            answer = f"Запомнил. ID: {mid}"
            await save_message(user_id, "assistant", answer)
            return answer
        return None

    if text.lower().startswith("заметка "):
        content = text.split(" ", 1)[1].strip()
        if content:
            nid = await add_note(user_id, title=content[:40], content=content)
            answer = f"Заметка #{nid} сохранена."
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
        answer = "Нет ответа."

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
    await update_behavior_pattern(user_id, text[:50])

    answer = await generate_jarvis_reply(user_id, text, save_user_message=False)
    if answer is None:
        return

    if not answer.startswith("Ошибка Groq") and not answer.startswith("Groq API временно"):
        proactive = await build_proactive_advice(user_id)
        if proactive:
            answer = f"{answer}\n\n{proactive}"

    await send_text_and_optional_voice(message, answer, settings)


# =========================
# COMMANDS — основные
# =========================

@router.message(Command("start"))
async def cmd_start(message: Message) -> None:
    user_id = message.from_user.id
    await ensure_user(user_id)
    settings = await get_settings(user_id)
    text = (
        "⚡ JARVIS ONLINE\n\n"
        "🧠 Интеллект: диалог, память, самообучение\n"
        "🛡 Безопасность: анализ угроз, тактика, риски\n"
        "🔬 Наука: физика, химия, математика, вирусология\n"
        "💊 Медицина: база знаний, мониторинг здоровья\n"
        "🌐 Разведка: веб-поиск, новости, соцсети\n"
        "📷 Анализ: фото, химический анализ, лица\n"
        "🗺 Навигация: маршруты, координаты\n"
        "💰 Финансы: анализ, прогноз бюджета\n"
        "📊 Аналитика: статистика, прогнозирование\n"
        "🔧 Самодиагностика и авторемонт\n\n"
        "Команды: /help\n"
        "Просто пиши — сам разберусь."
    )
    await message.answer(text, reply_markup=build_main_keyboard(settings))


@router.message(Command("help"))
async def cmd_help(message: Message) -> None:
    text = (
        "📋 КОМАНДЫ JARVIS\n\n"
        "🧠 ПАМЯТЬ\n"
        "/remember текст — сохранить\n"
        "/memory — показать\n"
        "/forget ID — удалить\n\n"
        "📌 ЗАДАЧИ\n"
        "/task текст — добавить\n"
        "/tasks — список\n"
        "/done ID — завершить\n\n"
        "⏰ НАПОМИНАНИЯ\n"
        "/remind 10m текст\n\n"
        "🛡 АНАЛИЗ\n"
        "/threat описание — анализ угрозы\n"
        "/winprob силы -> слабости — вероятность победы\n"
        "/tactic ситуация — тактика\n"
        "/behavior описание — анализ поведения\n"
        "/social описание — соцсвязи\n"
        "/risk сценарий — оценка риска\n\n"
        "🔬 НАУКА\n"
        "/calc выражение — калькулятор\n"
        "/stats 1,2,3,4 — статистика\n"
        "/chem вещество — химия\n"
        "/physics явление — физика\n"
        "/virus название — анализ вируса\n"
        "/med запрос — медицина\n\n"
        "💰 ФИНАНСЫ\n"
        "/income 5000 зарплата — доход\n"
        "/expense 200 продукты — расход\n"
        "/finance — анализ\n"
        "/budget — прогноз\n\n"
        "📊 АНАЛИТИКА\n"
        "/report — полный отчёт\n"
        "/forecast 1,2,3 — прогноз ряда\n\n"
        "🌐 ПОИСК\n"
        "/search запрос — веб\n"
        "/news тема — новости\n"
        "/translate язык текст\n\n"
        "🗺 НАВИГАЦИЯ\n"
        "/route откуда -> куда\n\n"
        "⚙️ СИСТЕМА\n"
        "/status — самодиагностика\n"
        "/threats — журнал угроз\n"
        "/profile — настройки\n"
        "/reset — очистить историю\n"
        "/voice on|off\n"
        "/model fast|smart|vision\n"
        "/speak текст — озвучить\n"
        "/learn тема — обучение\n"
        "/brief — брифинг\n"
    )
    await message.answer(text)


@router.message(Command("profile"))
async def cmd_profile(message: Message) -> None:
    user_id = message.from_user.id
    settings = await get_settings(user_id)
    await message.answer(
        f"⚙️ Профиль:\n"
        f"• voice_mode: {settings['voice_mode']}\n"
        f"• model_mode: {settings['model_mode']}\n"
        f"• briefing: {settings['briefing_enabled']}\n"
        f"• city: {settings.get('default_city') or 'не задан'}\n"
    )


@router.message(Command("reset"))
async def cmd_reset(message: Message) -> None:
    await ask_confirmation(message, "Очистить историю?", "reset_history", {})


@router.message(Command("voice"))
async def cmd_voice(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip().lower()
    if arg not in {"on", "off"}:
        await message.answer("Используй: /voice on или /voice off")
        return
    new_value = 1 if arg == "on" else 0
    await ask_confirmation(message, f"Voice mode → {'ON' if new_value else 'OFF'}?", "set_voice", {"value": new_value})


@router.message(Command("model"))
async def cmd_model(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip().lower()
    if arg not in {"fast", "smart", "vision"}:
        await message.answer("Используй: /model fast|smart|vision")
        return
    await ask_confirmation(message, f"Model → {arg}?", "set_model", {"value": arg})


@router.message(Command("briefing"))
async def cmd_briefing(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip().lower()
    if arg not in {"on", "off"}:
        await message.answer("Используй: /briefing on|off")
        return
    await ask_confirmation(message, f"Briefing → {arg.upper()}?", "set_briefing", {"value": 1 if arg == "on" else 0})


@router.message(Command("brief"))
async def cmd_brief(message: Message) -> None:
    await message.answer(await build_briefing(message.from_user.id))


@router.message(Command("city"))
async def cmd_city(message: Message, command: CommandObject) -> None:
    city = (command.args or "").strip()
    if not city:
        await message.answer("Используй: /city Ханчжоу")
        return
    await ask_confirmation(message, f"Город → «{city}»?", "set_city", {"city": city})


@router.message(Command("weather"))
async def cmd_weather(message: Message, command: CommandObject) -> None:
    city = (command.args or "").strip()
    if not city:
        settings = await get_settings(message.from_user.id)
        city = settings.get("default_city") or ""
    if not city:
        await message.answer("Укажи город: /weather Москва")
        return
    await message.answer(await get_weather_text(city))


@router.message(Command("speak"))
async def cmd_speak(message: Message, command: CommandObject) -> None:
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /speak текст")
        return
    audio_path = await tts_to_mp3(text)
    if not audio_path:
        await message.answer("Не смог озвучить.")
        return
    try:
        await message.answer_audio(FSInputFile(audio_path), caption="Озвучка")
    finally:
        try:
            os.remove(audio_path)
        except Exception:
            pass


# =========================
# COMMANDS — память / заметки
# =========================

@router.message(Command("remember"))
async def cmd_remember(message: Message, command: CommandObject) -> None:
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /remember текст")
        return
    mid = await add_memory(message.from_user.id, text)
    await message.answer(f"Запомнил. ID: {mid}")


@router.message(Command("memory"))
async def cmd_memory(message: Message):
   
    cursor.execute("SELECT id, content FROM memory ORDER BY id DESC LIMIT 30")
    rows = cursor.fetchall()
    
    rows = list(reversed(rows))
    
  
    rows_data = [f"• #{r['id']} {r['content']}" for r in rows]
    
    if not rows_data:
        await message.answer("🧠 Память пуста.")
        return

    current_chunk = "🧠 Память (последние 30 записей):\n"
    for line in rows_data:
      
        if len(current_chunk) + len(line) + 1 > 4000:
            await message.answer(current_chunk)
            current_chunk = ""
        current_chunk += line + "\n"
    
    if current_chunk:
        await message.answer(current_chunk)


@router.message(Command("forget"))
async def cmd_forget(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip()
    if not arg.isdigit():
        await message.answer("Используй: /forget ID")
        return
    await ask_confirmation(message, f"Удалить память #{arg}?", "delete_memory", {"memory_id": int(arg)})


@router.message(Command("note"))
async def cmd_note(message: Message, command: CommandObject) -> None:
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /note текст")
        return
    nid = await add_note(message.from_user.id, text[:40], text)
    await message.answer(f"Заметка #{nid} сохранена.")


@router.message(Command("notes"))
async def cmd_notes(message: Message) -> None:
    rows = await get_notes(message.from_user.id, limit=20)
    if not rows:
        await message.answer("Заметок нет.")
        return
    await message.answer("📝 Заметки:\n" + "\n".join([f"• #{r['id']} {r['title']} — {r['content'][:100]}" for r in rows]))


# =========================
# COMMANDS — задачи
# =========================

@router.message(Command("task"))
async def cmd_task(message: Message, command: CommandObject) -> None:
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /task текст [высокий/средний/низкий]")
        return
    priority = 2
    for kw, p in PRIORITY_MAP.items():
        if kw in text.lower():
            priority = p
            break
    task_id = await add_task(message.from_user.id, text, priority=priority)
    await message.answer(f"✅ Задача #{task_id} добавлена (приоритет: {priority})")


@router.message(Command("tasks"))
async def cmd_tasks(message: Message) -> None:
    rows = await get_tasks(message.from_user.id, status="open", limit=20)
    if not rows:
        await message.answer("Задач нет.")
        return
    icons = {1: "🔴", 2: "🟡", 3: "🟢"}
    lines = ["📌 Задачи:"]
    for r in rows:
        icon = icons.get(r.get("priority", 2), "⚪")
        due = f" — {r['due_at']}" if r.get("due_at") else ""
        lines.append(f"  {icon} #{r['id']} {r['text']}{due}")
    await message.answer("\n".join(lines))


@router.message(Command("done"))
async def cmd_done(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip()
    if not arg.isdigit():
        await message.answer("Используй: /done ID")
        return
    ok = await mark_task_done(message.from_user.id, int(arg))
    await message.answer("✅ Задача закрыта." if ok else "Задача не найдена.")


# =========================
# COMMANDS — напоминания
# =========================

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
    rid = await add_reminder(user_id, message.chat.id, parsed["text"], due)
    await message.answer(f"⏰ Напоминание #{rid}\nВремя: {due}")


# =========================
# COMMANDS — поиск
# =========================

@router.message(Command("search"))
async def cmd_search(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /search запрос")
        return
    scope = "all"
    query = raw
    low = raw.lower()
    for prefix, sc in (("memory ", "memory"), ("mem ", "memory"), ("tasks ", "tasks"), ("docs ", "docs"), ("notes ", "notes")):
        if low.startswith(prefix):
            scope = sc
            query = raw[len(prefix):].strip()
            break
    if scope == "all" and len(raw) > 3:
        result = await web_search(query)
        await message.answer(result)
        return
    results = await search_knowledge(query, user_id, top_k=5, scope=scope)
    if not results:
        await message.answer("Ничего не нашёл.")
        return
    lines = [f"🔍 Результаты ({scope}):"]
    for r in results:
        lines.append(f"• [{r['source']} #{r['id']}] {r['text'][:200]} (score={r['score']:.2f})")
    await message.answer("\n".join(lines))


@router.message(Command("news"))
async def cmd_news(message: Message, command: CommandObject) -> None:
    query = (command.args or "").strip()
    result = await fetch_news(query)
    await message.answer(result)


@router.message(Command("translate"))
async def cmd_translate(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /translate английский Привет мир")
        return
    parts = raw.split(None, 1)
    if len(parts) < 2:
        await message.answer("Формат: /translate язык текст")
        return
    result = await translate_text(parts[1], parts[0])
    await message.answer(result)


# =========================
# COMMANDS — аналитика и безопасность
# =========================

@router.message(Command("threat"))
async def cmd_threat(message: Message, command: CommandObject) -> None:
    desc = (command.args or "").strip()
    if not desc:
        await message.answer("Используй: /threat описание угрозы")
        return
    result = await analyze_threat(desc, message.from_user.id)
    await message.answer(result)


@router.message(Command("winprob"))
async def cmd_winprob(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()

    if not raw:
        await message.answer(
            "Отправь: сильные стороны и слабости противника.\n"
            "Можно так:\n"
            "/winprob я сильный в Python -> он слаб в опыте\n"
            "или просто:\n"
            "/winprob я сильный в Python противник слаб в опыте"
        )
        return

    # пробуем разные разделители
    if "->" in raw:
        left, right = raw.split("->", 1)
    elif "против" in raw:
        left, right = raw.split("против", 1)
    else:
        # fallback: делим пополам
        words = raw.split()
        mid = len(words) // 2
        left = " ".join(words[:mid])
        right = " ".join(words[mid:])

    left = left.strip()
    right = right.strip()

    if not left or not right:
        await message.answer("Не удалось разобрать данные. Попробуй переформулировать.")
        return

    result = await calculate_win_probability(left, right)
    await message.answer(result)


@router.message(Command("tactic"))
async def cmd_tactic(message: Message, command: CommandObject) -> None:
    situation = (command.args or "").strip()
    if not situation:
        await message.answer("Используй: /tactic описание ситуации")
        return
    result = await tactical_recommendations(situation)
    await message.answer(result)


@router.message(Command("behavior"))
async def cmd_behavior(message: Message, command: CommandObject) -> None:
    desc = (command.args or "").strip()
    if not desc:
        await message.answer("Используй: /behavior описание поведения человека")
        return
    result = await analyze_behavior(desc, message.from_user.id)
    await message.answer(result)


@router.message(Command("social"))
async def cmd_social(message: Message, command: CommandObject) -> None:
    desc = (command.args or "").strip()
    if not desc:
        await message.answer("Используй: /social описание связей")
        return
    result = await analyze_social_connections(desc)
    await message.answer(result)


@router.message(Command("risk"))
async def cmd_risk(message: Message, command: CommandObject) -> None:
    scenario = (command.args or "").strip()
    if not scenario:
        await message.answer("Используй: /risk сценарий")
        return
    result = await analyze_risk(scenario)
    await message.answer(result)


@router.message(Command("threats"))
async def cmd_threats(message: Message) -> None:
    user_id = message.from_user.id
    rows = await db_query(
        "SELECT threat_type, description, severity, resolved, created_at FROM threat_log WHERE user_id = ? ORDER BY id DESC LIMIT 15",
        (user_id,),
    )
    if not rows:
        await message.answer("🛡 Угроз не зафиксировано.")
        return
    lines = ["🛡 Журнал угроз:"]
    for r in rows:
        status = "✅" if r["resolved"] else "🔴"
        lines.append(f"{status} [{r['severity']}/10] {r['threat_type']}: {r['description'][:80]} ({r['created_at'][:10]})")
    await message.answer("\n".join(lines))


@router.message(Command("report"))
async def cmd_report(message: Message) -> None:
    result = await generate_analytical_report(message.from_user.id)
    await message.answer(result)


# =========================
# COMMANDS — наука
# =========================

@router.message(Command("calc"))
async def cmd_calc(message: Message, command: CommandObject) -> None:
    expr = (command.args or "").strip()
    if not expr:
        await message.answer("Используй: /calc 2+2*10 или /calc sqrt(144)")
        return
    await message.answer(calculate_math(expr))


@router.message(Command("stats"))
async def cmd_stats(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /stats 1,2,3,4,5")
        return
    try:
        data = [float(x.strip()) for x in re.split(r"[,;\s]+", raw) if x.strip()]
        if not data:
            raise ValueError
        result = calculate_statistics(data)
        lines = ["📊 Статистика:"]
        for k, v in result.items():
            lines.append(f"  {k}: {v}")
        await message.answer("\n".join(lines))
    except Exception:
        await message.answer("Не удалось разобрать данные. Пример: /stats 1,2,3,4,5")


@router.message(Command("forecast"))
async def cmd_forecast(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /forecast 10,20,30,40,50")
        return
    try:
        data = [float(x.strip()) for x in re.split(r"[,;\s]+", raw) if x.strip()]
        result = forecast_next(data, steps=3)
        if "error" in result:
            await message.answer(f"Ошибка: {result['error']}")
            return
        lines = [f"📈 Прогноз (следующие 3 периода):"]
        for i, v in enumerate(result["прогноз"], 1):
            lines.append(f"  +{i}: {v}")
        lines.append(f"Тренд: {result['тренд']} | R²: {result['R²']}")
        await message.answer("\n".join(lines))
    except Exception:
        await message.answer("Ошибка. Пример: /forecast 10,20,30,40,50")


@router.message(Command("chem"))
async def cmd_chem(message: Message, command: CommandObject) -> None:
    substance = (command.args or "").strip()
    if not substance:
        await message.answer("Используй: /chem H2O или /chem серная кислота")
        return
    result = await chemical_model(substance)
    for chunk in split_text(result):
        await message.answer(chunk)


@router.message(Command("physics"))
async def cmd_physics(message: Message, command: CommandObject) -> None:
    phenomenon = (command.args or "").strip()
    if not phenomenon:
        await message.answer("Используй: /physics свободное падение")
        return
    result = await physical_model(phenomenon)
    for chunk in split_text(result):
        await message.answer(chunk)


@router.message(Command("virus"))
async def cmd_virus(message: Message, command: CommandObject) -> None:
    name = (command.args or "").strip()
    if not name:
        await message.answer("Используй: /virus COVID-19")
        return
    result = await analyze_virus(name)
    for chunk in split_text(result):
        await message.answer(chunk)


@router.message(Command("med"))
async def cmd_med(message: Message, command: CommandObject) -> None:
    query = (command.args or "").strip()
    if not query:
        await message.answer("Используй: /med симптомы гриппа")
        return
    result = await medical_database_query(query)
    for chunk in split_text(result):
        await message.answer(chunk)


# =========================
# COMMANDS — навигация
# =========================

@router.message(Command("route"))
async def cmd_route(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw or "->" not in raw:
        await message.answer("Формат: /route Москва -> Санкт-Петербург")
        return
    parts = raw.split("->", 1)
    result = await build_route(parts[0].strip(), parts[1].strip())
    await message.answer(result)


# =========================
# COMMANDS — финансы
# =========================

@router.message(Command("income"))
async def cmd_income(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /income 5000 зарплата")
        return
    parts = raw.split(None, 1)
    try:
        amount = float(parts[0].replace(",", "."))
        category = parts[1].strip() if len(parts) > 1 else "прочее"
        rid = await add_finance_record(message.from_user.id, "income", amount, category)
        await message.answer(f"💚 Доход #{rid}: +{amount:,.0f} ({category})")
    except ValueError:
        await message.answer("Ошибка. Формат: /income 5000 зарплата")


@router.message(Command("expense"))
async def cmd_expense(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /expense 200 продукты")
        return
    parts = raw.split(None, 1)
    try:
        amount = float(parts[0].replace(",", "."))
        category = parts[1].strip() if len(parts) > 1 else "прочее"
        rid = await add_finance_record(message.from_user.id, "expense", amount, category)
        await message.answer(f"🔴 Расход #{rid}: -{amount:,.0f} ({category})")
    except ValueError:
        await message.answer("Ошибка. Формат: /expense 200 продукты")


@router.message(Command("finance"))
async def cmd_finance(message: Message) -> None:
    result = await finance_analysis(message.from_user.id)
    await message.answer(result)


@router.message(Command("budget"))
async def cmd_budget(message: Message, command: CommandObject) -> None:
    months = 3
    if command.args and command.args.strip().isdigit():
        months = min(int(command.args.strip()), 12)
    result = await forecast_budget(message.from_user.id, months)
    await message.answer(result)


# =========================
# COMMANDS — система
# =========================

@router.message(Command("status"))
async def cmd_status(message: Message) -> None:
    result = await self_repair_check()
    await message.answer(result)


@router.message(Command("learn"))
async def cmd_learn(message: Message, command: CommandObject) -> None:
    topic = (command.args or "").strip()
    if not topic:
        await message.answer("Используй: /learn Python или /learn квантовая физика")
        return
    level_map = {"начин": "начальный", "продв": "продвинутый", "сред": "средний"}
    level = "средний"
    for k, v in level_map.items():
        if k in topic.lower():
            level = v
            break
    result = await teach_topic(topic, level)
    for chunk in split_text(result):
        await message.answer(chunk)


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
    settings = await get_settings(callback.from_user.id)
    new_value = 0 if int(settings.get("voice_mode", 0)) else 1
    await ask_confirmation(callback.message, f"Voice mode → {'ON' if new_value else 'OFF'}?", "set_voice", {"value": new_value})
    await callback.answer()


@router.callback_query(F.data == "jarvis_briefing")
async def cb_briefing(callback: CallbackQuery) -> None:
    settings = await get_settings(callback.from_user.id)
    new_value = 0 if int(settings.get("briefing_enabled", 0)) else 1
    await ask_confirmation(callback.message, f"Briefing → {'ON' if new_value else 'OFF'}?", "set_briefing", {"value": new_value})
    await callback.answer()


@router.callback_query(F.data == "jarvis_weather")
async def cb_weather(callback: CallbackQuery) -> None:
    await callback.answer()
    settings = await get_settings(callback.from_user.id)
    city = settings.get("default_city")
    if city:
        await callback.message.answer(await get_weather_text(city))
    else:
        await callback.message.answer("Задай город: /city Москва")


@router.callback_query(F.data == "jarvis_tasks")
async def cb_tasks(callback: CallbackQuery) -> None:
    await callback.answer()
    rows = await get_tasks(callback.from_user.id, status="open", limit=10)
    if not rows:
        await callback.message.answer("Задач нет.")
        return
    icons = {1: "🔴", 2: "🟡", 3: "🟢"}
    lines = ["📌 Задачи:"] + [f"  {icons.get(r.get('priority',2), '⚪')} #{r['id']} {r['text']}" for r in rows]
    await callback.message.answer("\n".join(lines))


@router.callback_query(F.data == "jarvis_memory")
async def cb_memory(callback: CallbackQuery) -> None:
    await callback.answer()
    rows = await get_memories(callback.from_user.id, limit=10)
    if not rows:
        await callback.message.answer("Память пуста.")
        return
    await callback.message.answer("🧠 Память:\n" + "\n".join([f"• #{r['id']} {r['content']}" for r in rows]))


@router.callback_query(F.data == "jarvis_remind")
async def cb_remind(callback: CallbackQuery) -> None:
    await callback.answer()
    await callback.message.answer("⏰ Создай: /remind 10m текст")


@router.callback_query(F.data == "jarvis_threats")
async def cb_threats(callback: CallbackQuery) -> None:
    await callback.answer()
    rows = await db_query(
        "SELECT threat_type, description, severity, resolved FROM threat_log WHERE user_id = ? ORDER BY id DESC LIMIT 5",
        (callback.from_user.id,),
    )
    if not rows:
        await callback.message.answer("🛡 Угроз нет.")
        return
    lines = ["🛡 Последние угрозы:"]
    for r in rows:
        status = "✅" if r["resolved"] else "🔴"
        lines.append(f"{status} [{r['severity']}/10] {r['threat_type']}: {r['description'][:60]}")
    await callback.message.answer("\n".join(lines))


@router.callback_query(F.data == "jarvis_analytics")
async def cb_analytics(callback: CallbackQuery) -> None:
    await callback.answer()
    result = await generate_analytical_report(callback.from_user.id)
    await callback.message.answer(result)


@router.callback_query(F.data == "jarvis_finance")
async def cb_finance(callback: CallbackQuery) -> None:
    await callback.answer()
    result = await finance_analysis(callback.from_user.id)
    await callback.message.answer(result)


@router.callback_query(F.data == "jarvis_status")
async def cb_status(callback: CallbackQuery) -> None:
    await callback.answer()
    result = await self_repair_check()
    await callback.message.answer(result)


# =========================
# TEXT / VOICE / PHOTO / DOC
# =========================

@router.message(F.text)
async def on_text(message: Message) -> None:
    text = (message.text or "").strip()
    if not text or text.startswith("/"):
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
            await message.answer("Не смог распознать голосовое.")
            return
        lang = await detect_language(transcript)
        await message.answer(f"🎙 Распознал ({lang}):\n{transcript}")
        await process_user_text(message, transcript)
    except Exception as e:
        log_self_repair("on_voice", str(e), "exception handled")
        await message.answer("Ошибка при обработке голосового.")


@router.message(F.photo)
async def on_photo(message: Message) -> None:
    try:
        photo = message.photo[-1]
        tg_file = await bot.get_file(photo.file_id)
        buf = BytesIO()
        await bot.download_file(tg_file.file_path, destination=buf)
        image_bytes = buf.getvalue()
        caption = message.caption or ""

        # Определяем режим по подписи
        mode = "general"
        caption_lower = caption.lower()
        if any(kw in caption_lower for kw in ["химия", "химический", "состав", "вещество"]):
            mode = "chemical"
        elif any(kw in caption_lower for kw in ["угроза", "опасность", "безопасность"]):
            mode = "threat"
        elif any(kw in caption_lower for kw in ["лицо", "человек", "поведение", "эмоции"]):
            mode = "face"
        elif any(kw in caption_lower for kw in ["конструкция", "здание", "мост", "сооружение"]):
            mode = "construction"
        elif any(kw in caption_lower for kw in ["медицина", "анализ", "симптом", "рана"]):
            mode = "medical"

        result = await analyze_image_with_groq(image_bytes, caption=caption, mode=mode)
        await message.answer(result)

        # Проверяем на угрозы в результате анализа
        if "угроза" in result.lower() or "опасност" in result.lower():
            await log_threat(message.from_user.id, "photo_analysis", result[:300], severity=5)
    except Exception as e:
        log_self_repair("on_photo", str(e), "exception handled")
        await message.answer("Ошибка при анализе фото.")


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
            await message.answer("Поддерживаются: PDF, DOCX, XLSX, TXT, MD, CSV.")
            return
        if not text:
            await message.answer("Не смог извлечь текст.")
            return
        summary = await summarize_text(text, title=filename)
        await add_note(message.from_user.id, f"File: {filename}", summary)
        await add_document(message.from_user.id, filename, text, summary, ext)
        await auto_learn_from_text(message.from_user.id, text, source="document")
        if re.search(r"(медицин|анализ|врач|заключени|диагноз|обследован)", filename + " " + text, re.I):
            await add_memory(message.from_user.id, f"Медицинский документ: {filename}. {summary[:300]}")
        for chunk in split_text(f"📄 Файл обработан: {filename}\n\n{summary}"):
            await message.answer(chunk)
    except Exception as e:
        log_self_repair("on_document", str(e), "exception handled")
        await message.answer("Ошибка при обработке документа.")


# =========================
# SCHEDULER JOBS
# =========================

async def reminders_job() -> None:
    for r in await get_due_reminders(limit=100):
        try:
            await bot.send_message(r["chat_id"], f"⏰ Напоминание #{r['id']}:\n{r['text']}\n\nВремя: {r['due_at']}")
            await mark_reminder_sent(r["id"])
        except Exception as e:
            logger.warning("Reminder #%s failed: %s", r["id"], e)


async def recurring_reminders_job() -> None:
    for r in await get_due_recurring_reminders(limit=100):
        try:
            rule = json.loads(r["rule"])
            await bot.send_message(r["chat_id"], f"⏰ Напоминание #{r['id']}:\n{r['text']}")
            next_due = next_recurring_due(rule, datetime.now() + timedelta(minutes=1))
            if next_due:
                await update_recurring_reminder_next_due(r["id"], next_due)
        except Exception as e:
            logger.warning("Recurring reminder #%s failed: %s", r["id"], e)


async def briefing_job() -> None:
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M")
    for u in await db_query("SELECT user_id, briefing_time, last_brief_date FROM settings WHERE briefing_enabled = 1"):
        try:
            if u["briefing_time"] != current_time or u.get("last_brief_date") == today:
                continue
            await bot.send_message(u["user_id"], await build_briefing(u["user_id"]))
            await set_setting(u["user_id"], "last_brief_date", today)
        except Exception as e:
            logger.warning("Briefing failed for %s: %s", u["user_id"], e)


async def auto_self_repair_job() -> None:
    """Автоматический самоконтроль каждые 30 минут"""
    try:
        await db_query("SELECT COUNT(*) FROM settings")
    except Exception as e:
        log_self_repair("auto_repair_job", f"DB check failed: {e}", "reinit")
        init_db()

    global HTTP_SESSION
    if HTTP_SESSION and HTTP_SESSION.closed:
        HTTP_SESSION = None
        log_self_repair("auto_repair_job", "HTTP session was closed", "will recreate on next request")

    # Чистка старых подтверждений
    try:
        now_iso = datetime.now().isoformat(timespec="minutes")
        await db_execute("DELETE FROM confirmations WHERE expires_at < ?", (now_iso,))
    except Exception as e:
        log_self_repair("auto_repair_job", f"cleanup failed: {e}", "skipped")


async def health_monitor_job() -> None:
    """Мониторинг здоровья пользователей"""
    users = await db_query("SELECT DISTINCT user_id FROM health_metrics WHERE created_at >= datetime('now', '-1 day')")
    for u in users:
        try:
            user_id = u["user_id"]
            recent = await db_query(
                "SELECT metric, value FROM health_metrics WHERE user_id = ? AND created_at >= datetime('now', '-1 day') ORDER BY id DESC LIMIT 10",
                (user_id,),
            )
            metrics = {r["metric"]: r["value"] for r in recent if r["value"] is not None}
            warnings = await check_health_dangers(metrics)
            for w in warnings:
                existing = await db_query(
                    "SELECT id FROM threat_log WHERE user_id = ? AND description = ? AND created_at >= datetime('now', '-6 hour')",
                    (user_id, w[:300]),
                )
                if not existing:
                    await log_threat(user_id, "health_alert", w[:300], severity=7)
                    await bot.send_message(user_id, f"⚕️ ВНИМАНИЕ О ЗДОРОВЬЕ:\n{w}")
        except Exception as e:
            logger.warning("Health monitor failed for user: %s", e)


# =========================
# HTTP API
# =========================

async def handle_jarvis_request(request: web.Request) -> web.Response:
    if not JARVIS_API_KEY:
        return web.json_response({"error": "JARVIS_API_KEY не задан."}, status=503)
    if request.headers.get("X-API-Key", "") != JARVIS_API_KEY:
        return web.json_response({"error": "unauthorized"}, status=401)
    try:
        data = await request.json()
    except Exception:
        return web.json_response({"error": "invalid JSON"}, status=400)
    text = (data.get("text") or "").strip()
    if not text:
        return web.json_response({"error": "text required"}, status=400)
    user_id = JARVIS_HTTP_USER_ID
    if not user_id:
        return web.json_response({"error": "JARVIS_HTTP_USER_ID не задан."}, status=503)
    try:
        answer = await generate_jarvis_reply(user_id, text, save_user_message=True)
    except Exception as e:
        logger.exception("HTTP /jarvis error: %s", e)
        return web.json_response({"error": "internal error"}, status=500)
    return web.json_response({"reply": answer or "Готово."})


async def handle_health_check(request: web.Request) -> web.Response:
    return web.json_response({"status": "ok", "time": now_str(), "repairs": len(SELF_REPAIR_LOG)})


def build_http_app() -> web.Application:
    app = web.Application()
    app.router.add_post("/jarvis", handle_jarvis_request)
    app.router.add_get("/health", handle_health_check)
    return app


async def start_http_server() -> web.AppRunner:
    port = int(os.environ.get("PORT", "8080"))
    app = build_http_app()
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, host="0.0.0.0", port=port)
    await site.start()
    logger.info("HTTP API на порту %s", port)
    return runner


# =========================
# STARTUP / SHUTDOWN
# =========================

async def on_startup() -> None:
    init_db()
    scheduler.add_job(reminders_job, "interval", seconds=30, id="reminders", replace_existing=True)
    scheduler.add_job(recurring_reminders_job, "interval", seconds=30, id="recurring", replace_existing=True)
    scheduler.add_job(briefing_job, "interval", seconds=60, id="briefing", replace_existing=True)
    scheduler.add_job(auto_self_repair_job, "interval", minutes=30, id="self_repair", replace_existing=True)
    scheduler.add_job(health_monitor_job, "interval", minutes=60, id="health_monitor", replace_existing=True)
    scheduler.start()
    logger.info("JARVIS started. Self-repair: ACTIVE. Health monitor: ACTIVE.")


async def on_shutdown() -> None:
    try:
        scheduler.shutdown(wait=False)
    except Exception:
        pass
    global HTTP_SESSION
    if HTTP_SESSION and not HTTP_SESSION.closed:
        await HTTP_SESSION.close()
    await bot.session.close()
    logger.info("JARVIS stopped.")


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
