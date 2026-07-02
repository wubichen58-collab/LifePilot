"""
J.A.R.V.I.S. — Полная рабочая версия
═══════════════════════════════════════════════
Переменные окружения:
  BOT_TOKEN, GROQ_API_KEY, QWEATHER_KEY
  JARVIS_API_KEY, JARVIS_HTTP_USER_ID (опционально)
  PORT (default 8080)
"""

import asyncio
import base64
import json
import logging
import math
import os
import re
import sqlite3
import statistics
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

# ══════════════════════════════════════════════════════════════
# КОНФИГ
# ══════════════════════════════════════════════════════════════
BOT_TOKEN        = os.environ.get("BOT_TOKEN")
GROQ_API_KEY     = os.environ.get("GROQ_API_KEY")
QWEATHER_KEY     = os.environ.get("QWEATHER_KEY")
JARVIS_API_KEY   = os.environ.get("JARVIS_API_KEY")
JARVIS_HTTP_USER_ID = int(os.environ.get("JARVIS_HTTP_USER_ID", "0"))

GROQ_CHAT_URL      = "https://api.groq.com/openai/v1/chat/completions"
GROQ_TRANSCRIBE_URL = "https://api.groq.com/openai/v1/audio/transcriptions"

GROQ_MODEL_TEXT   = "llama-3.3-70b-versatile"
GROQ_MODEL_FAST   = "llama-3.1-8b-instant"
GROQ_MODEL_VISION = "meta-llama/llama-4-scout-17b-16e-instruct"
GROQ_MODEL_ASR    = "whisper-large-v3"

VOICE_NAME = "ru-RU-DmitryNeural"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH  = os.path.join(DATA_DIR, "jarvis.db")
os.makedirs(DATA_DIR, exist_ok=True)

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is required")

# ══════════════════════════════════════════════════════════════
# ЛОГИРОВАНИЕ
# ══════════════════════════════════════════════════════════════
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("jarvis")

# ══════════════════════════════════════════════════════════════
# БОТ
# ══════════════════════════════════════════════════════════════
bot       = Bot(token=BOT_TOKEN)
dp        = Dispatcher(storage=MemoryStorage())
router    = Router()
dp.include_router(router)
scheduler = AsyncIOScheduler()
DB_LOCK   = asyncio.Lock()
HTTP_SESSION: Optional[aiohttp.ClientSession] = None

# ══════════════════════════════════════════════════════════════
# БАЗА ДАННЫХ
# ══════════════════════════════════════════════════════════════
def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.executescript("""
        CREATE TABLE IF NOT EXISTS settings (
            user_id INTEGER PRIMARY KEY,
            voice_mode INTEGER DEFAULT 0,
            model_mode TEXT DEFAULT 'smart',
            briefing_enabled INTEGER DEFAULT 0,
            briefing_time TEXT DEFAULT '09:00',
            default_city TEXT DEFAULT 'Ханчжоу',
            home_address TEXT DEFAULT NULL,
            last_brief_date TEXT DEFAULT NULL,
            user_timezone TEXT DEFAULT 'Asia/Shanghai'
        );
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS personal_memory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            content TEXT NOT NULL,
            category TEXT DEFAULT 'general',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            text TEXT NOT NULL,
            status TEXT DEFAULT 'open',
            priority INTEGER DEFAULT 2,
            due_at TEXT DEFAULT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS reminders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            chat_id INTEGER NOT NULL,
            text TEXT NOT NULL,
            due_at TEXT NOT NULL,
            sent INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS health_stats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            metric TEXT NOT NULL,
            value REAL DEFAULT NULL,
            note TEXT DEFAULT NULL,
            recorded_date TEXT DEFAULT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS finance_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            record_type TEXT NOT NULL,
            amount REAL NOT NULL,
            category TEXT DEFAULT 'прочее',
            note TEXT DEFAULT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS threat_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            threat_type TEXT NOT NULL,
            description TEXT NOT NULL,
            severity INTEGER DEFAULT 1,
            resolved INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)
        conn.commit()

async def db_execute(sql: str, params: tuple = ()) -> int:
    async with DB_LOCK:
        def _run():
            with sqlite3.connect(DB_PATH) as conn:
                conn.row_factory = sqlite3.Row
                cur = conn.execute(sql, params)
                conn.commit()
                return int(cur.lastrowid or 0)
        return await asyncio.to_thread(_run)

async def db_rowcount(sql: str, params: tuple = ()) -> int:
    async with DB_LOCK:
        def _run():
            with sqlite3.connect(DB_PATH) as conn:
                cur = conn.execute(sql, params)
                conn.commit()
                return int(cur.rowcount)
        return await asyncio.to_thread(_run)

async def db_query(sql: str, params: tuple = ()) -> List[Dict]:
    async with DB_LOCK:
        def _run():
            with sqlite3.connect(DB_PATH) as conn:
                conn.row_factory = sqlite3.Row
                cur = conn.execute(sql, params)
                return [dict(r) for r in cur.fetchall()]
        return await asyncio.to_thread(_run)

async def db_one(sql: str, params: tuple = ()) -> Optional[Dict]:
    rows = await db_query(sql, params)
    return rows[0] if rows else None

# ══════════════════════════════════════════════════════════════
# НАСТРОЙКИ ПОЛЬЗОВАТЕЛЯ
# ══════════════════════════════════════════════════════════════
async def ensure_user(user_id: int) -> None:
    await db_execute("INSERT OR IGNORE INTO settings (user_id) VALUES (?)", (user_id,))

async def get_settings(user_id: int) -> Dict:
    await ensure_user(user_id)
    row = await db_one("SELECT * FROM settings WHERE user_id=?", (user_id,))
    return row or {"user_id": user_id, "voice_mode": 0, "model_mode": "smart",
                   "default_city": "Ханчжоу", "home_address": None}

async def set_setting(user_id: int, key: str, value: Any) -> None:
    await ensure_user(user_id)
    await db_execute(f"UPDATE settings SET {key}=? WHERE user_id=?", (value, user_id))

# ══════════════════════════════════════════════════════════════
# ИСТОРИЯ СООБЩЕНИЙ
# ══════════════════════════════════════════════════════════════
async def save_message(user_id: int, role: str, content: str) -> None:
    await db_execute(
        "INSERT INTO messages (user_id, role, content) VALUES (?,?,?)",
        (user_id, role, content)
    )

async def get_history(user_id: int, limit: int = 8) -> List[Dict]:
    rows = await db_query(
        "SELECT role, content FROM messages WHERE user_id=? ORDER BY id DESC LIMIT ?",
        (user_id, limit)
    )
    rows.reverse()
    return [{"role": r["role"], "content": r["content"]} for r in rows]

async def clear_history(user_id: int) -> None:
    await db_execute("DELETE FROM messages WHERE user_id=?", (user_id,))

# ══════════════════════════════════════════════════════════════
# ЛИЧНАЯ ПАМЯТЬ
# ══════════════════════════════════════════════════════════════
async def memory_add(user_id: int, content: str, category: str = "general") -> int:
    return await db_execute(
        "INSERT INTO personal_memory (user_id, content, category) VALUES (?,?,?)",
        (user_id, content, category)
    )

async def memory_list(user_id: int, limit: int = 30) -> List[Dict]:
    return await db_query(
        "SELECT id, content, category, created_at FROM personal_memory WHERE user_id=? ORDER BY id DESC LIMIT ?",
        (user_id, limit)
    )

async def memory_delete(user_id: int, mid: int) -> bool:
    n = await db_rowcount(
        "DELETE FROM personal_memory WHERE user_id=? AND id=?", (user_id, mid)
    )
    return n > 0

async def memory_get_context(user_id: int) -> str:
    rows = await db_query(
        "SELECT content FROM personal_memory WHERE user_id=? ORDER BY id DESC LIMIT 15",
        (user_id,)
    )
    if not rows:
        return ""
    return "Личная память:\n" + "\n".join(f"• {r['content']}" for r in rows)

# ══════════════════════════════════════════════════════════════
# ЗАДАЧИ
# ══════════════════════════════════════════════════════════════
async def task_add(user_id: int, text: str, priority: int = 2, due_at: str = None) -> int:
    return await db_execute(
        "INSERT INTO tasks (user_id, text, priority, due_at) VALUES (?,?,?,?)",
        (user_id, text, priority, due_at)
    )

async def task_list(user_id: int, limit: int = 20) -> List[Dict]:
    return await db_query(
        "SELECT id, text, priority, due_at FROM tasks WHERE user_id=? AND status='open' ORDER BY priority, id DESC LIMIT ?",
        (user_id, limit)
    )

async def task_done(user_id: int, tid: int) -> bool:
    n = await db_rowcount(
        "UPDATE tasks SET status='done' WHERE user_id=? AND id=?", (user_id, tid)
    )
    return n > 0

# ══════════════════════════════════════════════════════════════
# НАПОМИНАНИЯ
# ══════════════════════════════════════════════════════════════
async def reminder_add(user_id: int, chat_id: int, text: str, due_at: str) -> int:
    return await db_execute(
        "INSERT INTO reminders (user_id, chat_id, text, due_at) VALUES (?,?,?,?)",
        (user_id, chat_id, text, due_at)
    )

async def reminder_get_due() -> List[Dict]:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    return await db_query(
        "SELECT id, user_id, chat_id, text, due_at FROM reminders WHERE sent=0 AND due_at<=? ORDER BY due_at ASC LIMIT 100",
        (now,)
    )

async def reminder_mark_sent(rid: int) -> None:
    await db_execute("UPDATE reminders SET sent=1 WHERE id=?", (rid,))

# ══════════════════════════════════════════════════════════════
# ЗДОРОВЬЕ
# ══════════════════════════════════════════════════════════════
async def health_add(user_id: int, metric: str, value: float, note: str = "") -> int:
    today = datetime.now().strftime("%Y-%m-%d")
    return await db_execute(
        "INSERT INTO health_stats (user_id, metric, value, note, recorded_date) VALUES (?,?,?,?,?)",
        (user_id, metric, value, note, today)
    )

async def health_summary(user_id: int) -> str:
    rows = await db_query(
        "SELECT metric, AVG(value) as avg_v, MAX(value) as max_v, MIN(value) as min_v, COUNT(*) as cnt "
        "FROM health_stats WHERE user_id=? AND created_at >= datetime('now','-14 day') GROUP BY metric",
        (user_id,)
    )
    if not rows:
        return "Данных о здоровье нет. Заполни через /health"
    lines = ["💊 Статистика здоровья (14 дней):"]
    for r in rows:
        lines.append(f"• {r['metric']}: среднее {r['avg_v']:.1f} | мин {r['min_v']:.1f} | макс {r['max_v']:.1f} ({r['cnt']} записей)")
    return "\n".join(lines)

# ══════════════════════════════════════════════════════════════
# ФИНАНСЫ
# ══════════════════════════════════════════════════════════════
async def finance_add(user_id: int, rtype: str, amount: float, category: str, note: str = "") -> int:
    return await db_execute(
        "INSERT INTO finance_records (user_id, record_type, amount, category, note) VALUES (?,?,?,?,?)",
        (user_id, rtype, amount, category, note)
    )

async def finance_summary(user_id: int) -> str:
    rows = await db_query(
        "SELECT record_type, SUM(amount) as total, category FROM finance_records "
        "WHERE user_id=? GROUP BY record_type, category ORDER BY total DESC",
        (user_id,)
    )
    if not rows:
        return "Финансовых данных нет."
    income = sum(r["total"] for r in rows if r["record_type"] == "income")
    expense = sum(r["total"] for r in rows if r["record_type"] == "expense")
    balance = income - expense
    lines = [
        f"💰 Финансы:",
        f"• Доходы: {income:,.0f}",
        f"• Расходы: {expense:,.0f}",
        f"• Баланс: {balance:,.0f} {'✅' if balance >= 0 else '🔴'}",
    ]
    top_exp = [r for r in rows if r["record_type"] == "expense"][:5]
    if top_exp:
        lines.append("Топ расходов:")
        for r in top_exp:
            lines.append(f"  • {r['category']}: {r['total']:,.0f}")
    return "\n".join(lines)

# ══════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ══════════════════════════════════════════════════════════════
def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")

def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()

def split_text(text: str, limit: int = 3500) -> List[str]:
    if len(text) <= limit:
        return [text]
    parts, current = [], ""
    for line in text.split("\n"):
        candidate = (current + "\n" + line).strip()
        if len(candidate) <= limit:
            current = candidate
        else:
            if current:
                parts.append(current)
            current = line
    if current:
        parts.append(current)
    return [p for p in parts if p.strip()] or [text[:limit]]

def select_model(settings: Dict) -> str:
    mode = str(settings.get("model_mode", "smart")).lower()
    return {"fast": GROQ_MODEL_FAST, "vision": GROQ_MODEL_VISION}.get(mode, GROQ_MODEL_TEXT)

async def get_http() -> aiohttp.ClientSession:
    global HTTP_SESSION
    if HTTP_SESSION is None or HTTP_SESSION.closed:
        HTTP_SESSION = aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=30))
    return HTTP_SESSION

# ══════════════════════════════════════════════════════════════
# GROQ API
# ══════════════════════════════════════════════════════════════
async def groq_chat(messages: List[Dict], model: str, temperature: float = 0.5, max_tokens: int = 800) -> str:
    if not GROQ_API_KEY:
        return "GROQ_API_KEY не задан."
    session = await get_http()
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    payload = {"model": model, "messages": messages, "temperature": temperature, "max_tokens": max_tokens}
    for attempt in range(3):
        try:
            async with session.post(GROQ_CHAT_URL, headers=headers, json=payload) as resp:
                data = await resp.json(content_type=None)
                if resp.status == 429:
                    await asyncio.sleep(15 * (attempt + 1))
                    continue
                if resp.status >= 400:
                    return f"Ошибка API: {resp.status}"
                return data["choices"][0]["message"]["content"].strip()
        except Exception as e:
            await asyncio.sleep(3)
    return "Groq API недоступен."

async def groq_transcribe(audio_bytes: bytes, filename: str = "voice.ogg") -> Optional[str]:
    if not GROQ_API_KEY:
        return None
    session = await get_http()
    form = aiohttp.FormData()
    form.add_field("file", audio_bytes, filename=filename, content_type="audio/ogg")
    form.add_field("model", GROQ_MODEL_ASR)
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    try:
        async with session.post(GROQ_TRANSCRIBE_URL, headers=headers, data=form) as resp:
            data = await resp.json(content_type=None)
            return (data.get("text") or "").strip() or None
    except Exception as e:
        logger.error("Transcribe error: %s", e)
        return None

# ══════════════════════════════════════════════════════════════
# TTS
# ══════════════════════════════════════════════════════════════
async def tts_to_file(text: str) -> Optional[str]:
    text = normalize(text)
    if not text:
        return None
    fd, path = tempfile.mkstemp(suffix=".mp3", dir=DATA_DIR)
    os.close(fd)
    try:
        await edge_tts.Communicate(text[:1000], VOICE_NAME).save(path)
        return path
    except Exception as e:
        logger.error("TTS error: %s", e)
        try:
            os.remove(path)
        except Exception:
            pass
        return None

async def send_reply(message: Message, text: str, settings: Dict) -> None:
    for chunk in split_text(text):
        await message.answer(chunk)
    if int(settings.get("voice_mode", 0)):
        path = await tts_to_file(text[:800])
        if path:
            try:
                await message.answer_audio(FSInputFile(path))
            finally:
                try:
                    os.remove(path)
                except Exception:
                    pass

# ══════════════════════════════════════════════════════════════
# ПОГОДА (QWeather)
# ══════════════════════════════════════════════════════════════
CITY_LOCATION_CACHE: Dict[str, str] = {
    "ханчжоу": "101210101",
    "hangzhou": "101210101",
    "пекин": "101010100",
    "beijing": "101010100",
    "шанхай": "101020100",
    "shanghai": "101020100",
}

async def get_location_id(city: str) -> Optional[str]:
    city_lower = city.lower().strip()
    if city_lower in CITY_LOCATION_CACHE:
        return CITY_LOCATION_CACHE[city_lower]
    if not QWEATHER_KEY:
        return None
    session = await get_http()
    url = f"https://geoapi.qweather.com/v2/city/lookup?location={quote_plus(city)}&key={QWEATHER_KEY}&lang=ru"
    try:
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=8)) as resp:
            data = await resp.json(content_type=None)
            locs = data.get("location") or []
            if locs:
                loc_id = locs[0]["id"]
                CITY_LOCATION_CACHE[city_lower] = loc_id
                return loc_id
    except Exception as e:
        logger.error("City lookup error: %s", e)
    return None

async def get_weather(city: str) -> str:
    if not QWEATHER_KEY:
        return "QWEATHER_KEY не задан."
    loc_id = await get_location_id(city)
    if not loc_id:
        return f"Город '{city}' не найден."
    session = await get_http()
    url = f"https://devapi.qweather.com/v7/weather/now?location={loc_id}&key={QWEATHER_KEY}&lang=ru"
    try:
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=8)) as resp:
            data = await resp.json(content_type=None)
            if data.get("code") != "200":
                return f"Ошибка погоды: {data.get('code')}"
            now = data["now"]
            return (
                f"🌤 Погода в {city}:\n"
                f"• {now.get('text','—')} | {now.get('temp','—')}°C\n"
                f"• Ощущается: {now.get('feelsLike','—')}°C\n"
                f"• Влажность: {now.get('humidity','—')}%\n"
                f"• Ветер: {now.get('windDir','—')} {now.get('windSpeed','—')} км/ч"
            )
    except Exception as e:
        return f"Ошибка получения погоды: {e}"

# ══════════════════════════════════════════════════════════════
# ПАРСИНГ ВРЕМЕНИ И НАПОМИНАНИЙ
# ══════════════════════════════════════════════════════════════
def parse_due(expr: str) -> Optional[datetime]:
    expr = expr.strip().lower()

    # Относительное: 10m, 2h, 1d, 30s
    m = re.match(r"^(\d+)\s*([smhd])$", expr)
    if m:
        val = int(m.group(1))
        unit = m.group(2)
        delta = {"s": timedelta(seconds=val), "m": timedelta(minutes=val),
                 "h": timedelta(hours=val), "d": timedelta(days=val)}.get(unit)
        return datetime.now() + delta if delta else None

    # "через N минут/часов"
    m = re.search(r"через\s+(\d+)\s*(минут|мин|час|ч\b|секунд|сек|дн|день|дней)", expr)
    if m:
        val = int(m.group(1))
        unit = m.group(2)
        if "мин" in unit:
            return datetime.now() + timedelta(minutes=val)
        if "час" in unit or unit == "ч":
            return datetime.now() + timedelta(hours=val)
        if "сек" in unit:
            return datetime.now() + timedelta(seconds=val)
        if "дн" in unit or "день" in unit or "дней" in unit:
            return datetime.now() + timedelta(days=val)

    # "в HH:MM"
    m = re.search(r"в\s+(\d{1,2}):(\d{2})", expr)
    if m:
        dt = datetime.now().replace(hour=int(m.group(1)), minute=int(m.group(2)), second=0, microsecond=0)
        if dt <= datetime.now():
            dt += timedelta(days=1)
        return dt

    # "завтра в HH:MM" / "завтра"
    if "завтра" in expr:
        m2 = re.search(r"(\d{1,2}):(\d{2})", expr)
        base = (datetime.now() + timedelta(days=1)).replace(second=0, microsecond=0)
        if m2:
            return base.replace(hour=int(m2.group(1)), minute=int(m2.group(2)))
        return base.replace(hour=9, minute=0)

    # Абсолютные форматы
    for fmt in ("%Y-%m-%d %H:%M", "%d.%m.%Y %H:%M", "%d.%m %H:%M", "%d.%m.%Y", "%d.%m"):
        try:
            dt = datetime.strptime(expr, fmt)
            if fmt in ("%d.%m", "%d.%m.%Y"):
                dt = dt.replace(year=datetime.now().year, hour=9, minute=0)
            return dt
        except ValueError:
            pass

    return None

def detect_reminder_intent(text: str) -> Optional[Dict]:
    t = text.lower().strip()
    trigger_patterns = [
        r"(поставь|поставить|добавь|создай)\s+(таймер|напоминание|напомни)",
        r"напомни(шь)?\s+(мне\s+)?",
        r"через\s+\d+\s*(минут|мин|час|ч\b|секунд|сек)",
        r"таймер\s+на\s+\d+",
    ]
    if not any(re.search(p, t) for p in trigger_patterns):
        return None

    due = None
    m = re.search(r"через\s+(\d+)\s*(минут|мин|час|ч\b|секунд|сек|дн|день|дней)", t)
    if m:
        due = parse_due(m.group(0))

    if not due:
        m = re.search(r"на\s+(\d+)\s*(минут|мин|час|ч\b|секунд|сек)", t)
        if m:
            val = int(m.group(1))
            unit = m.group(2)
            if "мин" in unit:
                due = datetime.now() + timedelta(minutes=val)
            elif "час" in unit or unit == "ч":
                due = datetime.now() + timedelta(hours=val)
            elif "сек" in unit:
                due = datetime.now() + timedelta(seconds=val)

    if not due:
        m = re.search(r"в\s+(\d{1,2}):(\d{2})", t)
        if m:
            dt = datetime.now().replace(hour=int(m.group(1)), minute=int(m.group(2)), second=0, microsecond=0)
            if dt <= datetime.now():
                dt += timedelta(days=1)
            due = dt

    if not due:
        return None

    reminder_text = t
    for pat in [
        r"(поставь|поставить|добавь|создай)\s+(таймер|напоминание)\s*",
        r"напомни(шь)?\s+(мне\s+)?",
        r"через\s+\d+\s*(минут|мин|час|ч\b|секунд|сек)\s*",
        r"на\s+\d+\s*(минут|мин|час|ч\b|секунд|сек)\s*",
        r"в\s+\d{1,2}:\d{2}\s*",
        r"(завтра|сегодня)\s*",
        r"таймер\s*",
    ]:
        reminder_text = re.sub(pat, "", reminder_text).strip()

    reminder_text = reminder_text.strip(" ,.-") or "Таймер"
    return {"due": due, "text": reminder_text}

# ══════════════════════════════════════════════════════════════
# МАРШРУТЫ
# ══════════════════════════════════════════════════════════════
async def build_route(from_place: str, to_place: str, mode: str = "driving") -> str:
    mode_ru = {"driving": "на автомобиле", "walking": "пешком", "transit": "на общественном транспорте"}.get(mode, mode)
    system = (
        "Ты навигационный ассистент для Китая. Строй маршруты по Ханчжоу и другим городам Китая.\n"
        "Дай: пошаговый маршрут, примерное время, расстояние, транспортные советы.\n"
        "Учитывай китайские реалии: метро, автобусы, такси DiDi, велосипеды.\n"
        "Если маршрут внутри города — предложи несколько вариантов транспорта."
    )
    return await groq_chat(
        [
            {"role": "system", "content": system},
            {"role": "user", "content": f"Маршрут {mode_ru}: из «{from_place}» в «{to_place}»"}
        ],
        model=GROQ_MODEL_FAST, 
        temperature=0.3, 
        max_tokens=600
    )

def detect_route_intent(text: str) -> Optional[Dict]:
    t = text.lower()
    route_triggers = [
        r"(построй|проложи|покажи|как доехать|как дойти|как добраться)\s+(маршрут|путь|дорогу)?\s*",
        r"маршрут\s+(от|из|до|к|в)",
        r"(от меня|от дома|с моего адреса)\s+(до|к|в)",
    ]
    if not any(re.search(p, t) for p in route_triggers):
        return None

    from_place = None
    if re.search(r"(от меня|от дома|с моего адреса|из дома)", t):
        from_place = "HOME"

    to_match = re.search(r"(до|к|в|на)\s+([^,\n]+?)(?:\s+(?:пешком|на такси|на метро|на автобусе|на машине))?$", t)
    to_place = to_match.group(2).strip() if to_match else None

    mode = "driving"
    if "пешком" in t:
        mode = "walking"
    elif "метро" in t or "автобус" in t or "общественн" in t:
        mode = "transit"

    return {"from": from_place, "to": to_place, "mode": mode} if to_place else None

# ══════════════════════════════════════════════════════════════
# АНАЛИЗ ИЗОБРАЖЕНИЙ
# ══════════════════════════════════════════════════════════════
async def analyze_image(image_bytes: bytes, caption: str = "") -> str:
    if not GROQ_API_KEY:
        return "GROQ_API_KEY не задан."
    b64 = base64.b64encode(image_bytes).decode()
    system = (
        "Ты анализируешь изображения. Отвечай на русском.\n"
        "Опиши подробно что видишь: объекты, текст, люди, место, действия.\n"
        "Если есть текст на photo — извлеки его полностью.\n"
        "Если есть ошибки/предупреждения — объясни их.\n"
        "Если спрашивают конкретно — отвечай точно на вопрос."
    )
    user_text = caption if caption else "Подробно опиши что на этом изображении."
    payload = {
        "model": GROQ_MODEL_VISION,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": [
                {"type": "text", "text": user_text},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
            ]},
        ],
        "temperature": 0.2,
        "max_tokens": 800,
    }
    session = await get_http()
    try:
        async with session.post(GROQ_CHAT_URL,
                                headers={"Authorization": f"Bearer {GROQ_API_KEY}"},
                                json=payload,
                                timeout=aiohttp.ClientTimeout(total=30)) as resp:
            data = await resp.json(content_type=None)
            if resp.status >= 400:
                return f"Ошибка анализа: {resp.status}"
            return data["choices"][0]["message"]["content"].strip()
    except Exception as e:
        return f"Ошибка при анализе изображения: {e}"

# ══════════════════════════════════════════════════════════════
# НАУЧНЫЕ ВЫЧИСЛЕНИЯ
# ══════════════════════════════════════════════════════════════
def calc_math(expression: str) -> str:
    safe_globals = {
        "__builtins__": {}, "math": math,
        "sqrt": math.sqrt, "sin": math.sin, "cos": math.cos, "tan": math.tan,
        "log": math.log, "log10": math.log10, "exp": math.exp, "pi": math.pi,
        "e": math.e, "abs": abs, "round": round, "pow": pow, "factorial": math.factorial,
    }
    try:
        cleaned = expression.replace("^", "**")
        result = eval(cleaned, safe_globals)
        return f"Результат: {result}"
    except Exception as ex:
        return f"Ошибка вычисления: {ex}"

def calc_stats(data: List[float]) -> str:
    if not data:
        return "Нет данных."
    try:
        avg = statistics.mean(data)
        med = statistics.median(data)
        sd  = statistics.stdev(data) if len(data) > 1 else 0.0
        c   = Counter(data)
        mode_val = c.most_common(1)[0][0]
        return (
            f"📊 Статистика ({len(data)} чисел):\n"
            f"• Среднее: {avg:.4f}\n"
            f"• Медиана: {med:.4f}\n"
            f"• Мода: {mode_val}\n"
            f"• Ст. отклонение: {sd:.4f}\n"
            f"• Сумма: {sum(data):.4f}\n"
            f"• Мин/Макс: {min(data)} / {max(data)}"
        )
    except Exception as e:
        return f"Ошибка статистики: {e}"

# ══════════════════════════════════════════════════════════════
# ЯДРО ДИАЛОГА JARVIS
# ══════════════════════════════════════════════════════════════
async def build_system_prompt(user_id: int, current_text: str, settings: Dict) -> str:
    memory_ctx = await memory_get_context(user_id)
    return (
        "Ты J.A.R.V.I.S. — высокоинтеллектуальный ИИ-ассистент, созданный Тони Старком.\n"
        "Отвечай уверенно, кратко, по делу, с уважением и легким технологическим акцентом (используй имя сэр, если уместно).\n"
        f"Текущее время: {now_str()} (Базовый город: {settings.get('default_city','Ханчжоу')}).\n"
        f"{memory_ctx}\n"
        "Используй эти данные о пользователе для контекста, не переспрашивай то, что уже знаешь."
    )

async def jarvis_reply(user_id: int, text: str, settings: Dict, save: bool = True) -> str:
    text = normalize(text)
    if not text:
        return ""
    if save:
        await save_message(user_id, "user", text)
    system = await build_system_prompt(user_id, text, settings)
    history = await get_history(user_id, limit=8)
    messages = [{"role": "system", "content": system}] + history
    if not history or history[-1]["content"] != text:
        messages.append({"role": "user", "content": text})
    model = select_model(settings)
    answer = await groq_chat(messages, model=model, temperature=0.5, max_tokens=800)
    await save_message(user_id, "assistant", answer)
    return answer

# ══════════════════════════════════════════════════════════════
# ОБРАБОТКА НАМЕРЕНИЙ ИЗ ТЕКСТА
# ══════════════════════════════════════════════════════════════
async def handle_intents(message: Message, text: str, settings: Dict) -> bool:
    user_id = message.from_user.id
    t = normalize(text).lower()

    # ── Напоминание/таймер
    reminder = detect_reminder_intent(t)
    if reminder:
        due_str = reminder["due"].strftime("%Y-%m-%d %H:%M")
        rid = await reminder_add(user_id, message.chat.id, reminder["text"], due_str)
        await message.answer(
            f"⏰ Напоминание #{rid} установлено\n"
            f"Когда: {due_str}\n"
            f"Текст: {reminder['text']}"
        )
        return True

    # ── Маршрут
    route = detect_route_intent(t)
    if route:
        from_place = route["from"]
        to_place = route["to"]
        if from_place == "HOME":
            addr = settings.get("home_address")
            from_place = addr if addr else "Ханчжоу (мой адрес не задан, используй /address)"
        if to_place:
            await message.answer("🗺 Строю маршрут...")
            result = await build_route(from_place or "Ханчжоу", to_place, route["mode"])
            await send_reply(message, result, settings)
            return True

    # ── Погода
    if any(w in t for w in ["погода", "weather", "温度", "天气", "дождь", "температура", "холодно", "жарко"]):
        city = settings.get("default_city") or "Ханчжоу"
        words = t.split()
        for idx, w in enumerate(words):
            if w in ["в", "погода", "город"] and idx + 1 < len(words):
                cand = words[idx+1].strip("?,.!")
                if len(cand) > 2 and cand not in ["погода", "weather", "какая"]:
                    city = cand
                    break
        await message.answer("🌤 Запрашиваю метеосводку...")
        w_res = await get_weather(city)
        await send_reply(message, w_res, settings)
        return True

    return False

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — ОБЩИЕ И НАСТРОЙКИ
# ══════════════════════════════════════════════════════════════
@router.message(Command("start"))
async def cmd_start(message: Message) -> None:
    await ensure_user(message.from_user.id)
    await message.answer(
        "⚡ J.A.R.V.I.S. ОНЛАЙН\n\n"
        "Просто пиши или говори — сам пойму.\n\n"
        "Основные команды:\n"
        "/help — все команды\n"
        "/city [город] — установить город для погоды\n"
        "/address [адрес] — твой домашний адрес\n"
        "/memory — личная память\n"
        "/tasks — задачи\n"
        "/health — статистика здоровья\n"
        "/finance — финансы\n"
        "/brief — дневной брифинг\n"
        "/status — диагностика"
    )

@router.message(Command("help"))
async def cmd_help(message: Message) -> None:
    await message.answer(
        "📋 J.A.R.V.I.S. — Команды\n\n"
        "⚙️ НАСТРОЙКИ\n"
        "/city Ханчжоу — город для погоды\n"
        "/address Ханчжоу, ул. Xxx — домашний адрес\n"
        "/datetime 2024-12-31 09:00 — установить дату/время\n"
        "/voice on|off — голосовые ответы\n"
        "/model fast|smart — скорость vs качество\n"
        "/reset — очистить историю\n\n"
        "🧠 ПАМЯТЬ\n"
        "/mem — показать личную память\n"
        "/mem+ текст — добавить в память\n"
        "/mem- ID — удалить из памяти\n"
        "(или просто напиши: «запомни что...»)\n\n"
        "📌 ЗАДАЧИ\n"
        "/task текст — добавить задачу\n"
        "/tasks — список задач\n"
        "/done ID — завершить задачу\n\n"
        "⏰ НАПОМИНАНИЯ\n"
        "/remind 10m текст — добавить напоминание\n"
        "или просто: «напомни через 20 минут позвонить»\n"
        "или: «поставь таймер на 1 час»\n\n"
        "🌤 ПОГОДА\n"
        "/weather [город] — погода\n\n"
        "🗺 МАРШРУТЫ\n"
        "/route откуда -> куда — маршрут\n"
        "или: «как доехать до West Lake»\n"
        "или: «построй маршрут от меня до...»\n\n"
        "💊 ЗДОРОВЬЕ\n"
        "/health — показать статистику\n"
        "/health сон 7.5 — записать сон\n"
        "/health вес 72.5 — записать вес\n\n"
        "💰 ФИНАНСЫ\n"
        "/finance — отчет за месяц\n"
        "/income 5000 зарплата — доход\n"
        "/expense 150 кофе — расход\n\n"
        "🧮 НАУКА\n"
        "/calc 2+2*sqrt(16) — калькулятор\n"
        "/stats 10,23,42,11 — мат. статистика\n"
        "/winprob сила -> слабость — вероятность победы"
    )

@router.message(Command("city"))
async def cmd_city(message: Message, command: CommandObject) -> None:
    city = (command.args or "").strip()
    if not city:
        await message.answer("Используй: /city НазваниеГорода")
        return
    await set_setting(message.from_user.id, "default_city", city)
    await message.answer(f"✅ Город по умолчанию изменен на: {city}")

@router.message(Command("address"))
async def cmd_address(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    addr = (command.args or "").strip()
    if not addr:
        settings = await get_settings(user_id)
        current = settings.get("home_address") or "не задан"
        await message.answer(f"Твой адрес: {current}\nУстановить: /address Ханчжоу, район Xihu, ул. XXX")
        return
    await set_setting(user_id, "home_address", addr)
    await memory_add(user_id, f"Мой домашний адрес: {addr}", category="personal")
    await message.answer(f"✅ Адрес сохранён: {addr}")

@router.message(Command("datetime"))
async def cmd_datetime(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer(
            f"Текущее время бота: {now_str()} (UTC+8)\n\n"
            f"Используй: /datetime 2024-12-31 09:00 — я запомню это как ориентир."
        )
        return
    await memory_add(message.from_user.id, f"Пользователь указал текущее время: {raw}", category="system")
    await message.answer(f"✅ Запомнил: {raw}. Буду учитывать при ответах.")

@router.message(Command("voice"))
async def cmd_voice(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip().lower()
    if arg not in {"on", "off"}:
        await message.answer("Используй: /voice on или /voice off")
        return
    await set_setting(message.from_user.id, "voice_mode", 1 if arg == "on" else 0)
    await message.answer(f"🔊 Голос: {'включён' if arg == 'on' else 'выключен'}")

@router.message(Command("model"))
async def cmd_model(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip().lower()
    if arg not in {"fast", "smart"}:
        await message.answer("Используй: /model fast (быстро) или /model smart (умнее)")
        return
    await set_setting(message.from_user.id, "model_mode", arg)
    await message.answer(f"🧠 Модель: {arg.upper()}")

@router.message(Command("reset"))
async def cmd_reset(message: Message) -> None:
    await clear_history(message.from_user.id)
    await message.answer("🧹 История диалога очищена.")

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — БАЗА ЗНАНИЙ И ПАМЯТЬ
# ══════════════════════════════════════════════════════════════
@router.message(Command("mem"))
async def cmd_mem(message: Message, command: CommandObject) -> None:
    rows = await memory_list(message.from_user.id)
    if not rows:
        await message.answer("🧠 Память пуста. Скажи мне «Запомни, что...» или добавь через /mem+ текст")
        return
    lines = ["🧠 Личная память (последние 30 записей):"]
    for r in rows:
        lines.append(f"• `#{r['id']}` [{r['category']}]: {r['content']}")
    await message.answer("\n".join(lines))

@router.message(Command("mem+"))
async def cmd_mem_add(message: Message, command: CommandObject) -> None:
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /mem+ текст, который нужно запомнить")
        return
    mid = await memory_add(message.from_user.id, text)
    await message.answer(f"✅ Успешно записано в память под `#{mid}`: {text}")

@router.message(Command("mem-"))
async def cmd_mem_del(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip()
    if not arg.isdigit():
        await message.answer("Используй: /mem- ID (номер из /mem)")
        return
    ok = await memory_delete(message.from_user.id, int(arg))
    await message.answer(f"✅ Удалено #{arg}" if ok else f"❌ Запись #{arg} не найдена")

@router.message(Command("memory"))
async def cmd_memory_alias(message: Message, command: CommandObject) -> None:
    await cmd_mem(message, command)

@router.message(Command("remember"))
async def cmd_remember(message: Message, command: CommandObject) -> None:
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /remember текст")
        return
    mid = await memory_add(message.from_user.id, text)
    await message.answer(f"✅ Запомнил #{mid}")

@router.message(Command("forget"))
async def cmd_forget(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip()
    if not arg.isdigit():
        await message.answer("Используй: /forget ID")
        return
    ok = await memory_delete(message.from_user.id, int(arg))
    await message.answer(f"✅ Удалено" if ok else "❌ Не найдено")

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — ЗАДАЧИ
# ══════════════════════════════════════════════════════════════
PRIORITY_MAP = {"высокий": 1, "срочно": 1, "средний": 2, "низкий": 3}

@router.message(Command("task"))
async def cmd_task(message: Message, command: CommandObject) -> None:
    text = (command.args or "").strip()
    if not text:
        await message.answer("Используй: /task текст задачи")
        return
    priority = 2
    for kw, p in PRIORITY_MAP.items():
        if kw in text.lower():
            priority = p
            break
    tid = await task_add(message.from_user.id, text, priority)
    await message.answer(f"✅ Задача #{tid} добавлена")

@router.message(Command("tasks"))
async def cmd_tasks(message: Message) -> None:
    rows = await task_list(message.from_user.id)
    if not rows:
        await message.answer("📋 Задач нет.")
        return
    lines = ["📋 Открытые задачи (по приоритету):"]
    for r in rows:
        p_str = {1: "🔴 Срочно", 2: "🟡 Средний", 3: "🟢 Низкий"}.get(r["priority"], str(r["priority"]))
        lines.append(f"• `#{r['id']}` [{p_str}]: {r['text']}")
    await message.answer("\n".join(lines))

@router.message(Command("done"))
async def cmd_done(message: Message, command: CommandObject) -> None:
    arg = (command.args or "").strip()
    if not arg.isdigit():
        await message.answer("Используй: /done ID (номер задачи)")
        return
    ok = await task_done(message.from_user.id, int(arg))
    await message.answer(f"✅ Задача #{arg} отмечена выполненной!" if ok else f"❌ Задача #{arg} не найдена")

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — НАПОМИНАНИЯ
# ══════════════════════════════════════════════════════════════
@router.message(Command("remind"))
async def cmd_remind(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /remind 10m купить молоко")
        return

    parts = raw.split(None, 1)
    due = parse_due(parts[0])
    text = parts[1].strip() if len(parts) > 1 else "Напоминание"

    if not due and len(parts) >= 2:
        parts2 = raw.split(None, 2)
        if len(parts2) >= 2:
            due = parse_due(f"{parts2[0]} {parts2[1]}")
            text = parts2[2].strip() if len(parts2) > 2 else "Напоминание"

    if not due:
        await message.answer("Не понял формат времени.\nПример: /remind 10m текст или /remind завтра в 09:00 текст")
        return

    due_str = due.strftime("%Y-%m-%d %H:%M")
    rid = await reminder_add(user_id, message.chat.id, text, due_str)
    await message.answer(f"⏰ Напоминание #{rid}\nВремя: {due_str}\nТекст: {text}")

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — ПОГОДА И МАРШРУТЫ
# ══════════════════════════════════════════════════════════════
@router.message(Command("weather"))
async def cmd_weather(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    city = (command.args or "").strip()
    if not city:
        settings = await get_settings(user_id)
        city = settings.get("default_city") or "Ханчжоу"
    result = await get_weather(city)
    await message.answer(result)

@router.message(Command("route"))
async def cmd_route(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    raw = (command.args or "").strip()
    if not raw:
        await message.answer(
            "Примеры:\n"
            "/route West Lake -> Alibaba HQ\n"
            "/route от меня -> West Lake\n"
            "Или просто напиши: «как доехать до West Lake»"
        )
        return
    settings = await get_settings(user_id)
    if "->" in raw:
        parts = raw.split("->", 1)
        from_place = parts[0].strip()
        to_place = parts[1].strip()
    else:
        from_place = settings.get("default_city") or "Ханчжоу"
        to_place = raw

    if re.search(r"(от меня|от дома|мой адрес)", from_place.lower()):
        addr = settings.get("home_address")
        from_place = addr if addr else "Ханчжоу"

    await message.answer("🗺 Прокладываю маршрут...")
    result = await build_route(from_place, to_place, "driving")
    await send_reply(message, result, settings)

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — ЗДОРОВЬЕ
# ══════════════════════════════════════════════════════════════
@router.message(Command("health"))
async def cmd_health(message: Message, command: CommandObject) -> None:
    user_id = message.from_user.id
    raw = (command.args or "").strip()
    if not raw:
        await message.answer(await health_summary(user_id))
        return

    parts = raw.split(None, 2)
    metric = parts[0].lower()
    metric_ru = metric

    metrics_aliases = {
        "сон": "sleep", "sleep": "sleep",
        "вода": "water", "water": "water",
        "вес": "weight", "weight": "weight",
        "пульс": "pulse", "pulse": "pulse",
        "шаги": "steps", "steps": "steps",
        "калории": "calories", "calories": "calories",
    }
    if metric in metrics_aliases:
        metric = metrics_aliases[metric]

    try:
        value = float(parts[1].replace(",", "."))
    except (IndexError, ValueError):
        await message.answer(f"Не понял значение «{parts[1] if len(parts)>1 else ''}». Введи число.")
        return

    note = parts[2] if len(parts) > 2 else ""
    await health_add(user_id, metric, value, note)
    units = {"sleep": "ч", "water": "мл", "weight": "кг", "pulse": "уд/мин", "steps": "шагов", "calories": "ккал"}
    unit = units.get(metric, "")
    await message.answer(f"✅ Записано: {metric_ru} = {value}{unit}")

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — ФИНАНСЫ
# ══════════════════════════════════════════════════════════════
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
        rid = await finance_add(message.from_user.id, "income", amount, category)
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
        rid = await finance_add(message.from_user.id, "expense", amount, category)
        await message.answer(f"🔴 Расход #{rid}: -{amount:,.0f} ({category})")
    except ValueError:
        await message.answer("Ошибка. Формат: /expense 200 продукты")

@router.message(Command("finance"))
async def cmd_finance(message: Message) -> None:
    await message.answer(await finance_summary(message.from_user.id))

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — НАУКА И ВЫЧИСЛЕНИЯ
# ══════════════════════════════════════════════════════════════
@router.message(Command("calc"))
async def cmd_calc(message: Message, command: CommandObject) -> None:
    expr = (command.args or "").strip()
    if not expr:
        await message.answer("Используй: /calc 2 + 2")
        return
    await message.answer(calc_math(expr))

@router.message(Command("stats"))
async def cmd_stats(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /stats 1,2,3,4,5")
        return
    try:
        nums = [float(x.strip()) for x in re.split(r"[,\s]+", raw) if x.strip()]
        await message.answer(calc_stats(nums))
    except ValueError:
        await message.answer("Ошибка парсинга чисел. Вводи через запятую или пробел.")

@router.message(Command("winprob"))
async def cmd_winprob(message: Message, command: CommandObject) -> None:
    raw = (command.args or "").strip()
    if not raw:
        await message.answer("Используй: /winprob наши плюсы -> их минусы")
        return
    if "->" in raw:
        parts = raw.split("->", 1)
        prompt = f"Мои сильные стороны: {parts[0].strip()}\nСлабости противника: {parts[1].strip()}"
    else:
        prompt = raw
    system = "Ты военный аналитик. Рассчитай вероятность победы.\n• Анализ сторон\n• Вероятность: X%\n• Тактика и Риски"
    result = await groq_chat([{"role": "system", "content": system}, {"role": "user", "content": prompt}],
                             model=GROQ_MODEL_TEXT, temperature=0.3, max_tokens=600)
    await message.answer(result)

# ══════════════════════════════════════════════════════════════
# КОМАНДЫ — БРИФИНГ И СТАТУС
# ══════════════════════════════════════════════════════════════
@router.message(Command("brief"))
async def cmd_brief(message: Message) -> None:
    user_id = message.from_user.id
    settings = await get_settings(user_id)
    city = settings.get("default_city") or "Ханчжоу"
    weather = await get_weather(city)
    tasks = await task_list(user_id, limit=5)
    health = await health_summary(user_id)
    lines = [f"☀️ Брифинг — {now_str()}", "", weather, ""]
    if tasks:
        lines.append(f"📋 Задач: {len(tasks)}")
        for t in tasks[:3]:
            lines.append(f" • #{t['id']} {t['text']}")
    else:
        lines.append("📋 Задач нет.")
    lines.append("")
    lines.append(health)
    await message.answer("\n".join(lines))

@router.message(Command("status"))
async def cmd_status(message: Message) -> None:
    lines = [
        f"🔧 Статус J.A.R.V.I.S.",
        f"• Время: {now_str()}",
        f"• Groq API: Доступен",
        f"• Системные модули: Стабильно",
    ]
    await message.answer("\n".join(lines))

# ══════════════════════════════════════════════════════════════
# МЕДИА-ХЕНДЛЕРЫ И СВОБОДНЫЙ ТЕКСТ
# ══════════════════════════════════════════════════════════════
async def process_audio(audio_bytes: bytes, message: Message, filename: str = "voice.ogg") -> None:
    user_id = message.from_user.id
    settings = await get_settings(user_id)
    await message.answer("⚡ *Распознаю аудио поток...*")
    text = await groq_transcribe(audio_bytes, filename)
    if not text:
        await message.answer("❌ Мне не удалось разобрать слова. Попробуйте записать еще раз.")
        return
    await message.answer(f"🗣 *Вы:* {text}")
    if await handle_intents(message, text, settings):
        return
    reply = await jarvis_reply(user_id, text, settings)
    await send_reply(message, reply, settings)

@router.message(F.voice)
async def on_voice(message: Message) -> None:
    tg_file = await bot.get_file(message.voice.file_id)
    buf = BytesIO()
    await bot.download_file(tg_file.file_path, destination=buf)
    await process_audio(buf.getvalue(), message, "voice.ogg")

@router.message(F.video_note)
async def on_video_note(message: Message) -> None:
    await message.answer("🎥 Обрабатываю видеосообщение...")
    tg_file = await bot.get_file(message.video_note.file_id)
    buf = BytesIO()
    await bot.download_file(tg_file.file_path, destination=buf)
    await process_audio(buf.getvalue(), message, "video_note.mp4")

@router.message(F.photo)
async def on_photo(message: Message) -> None:
    user_id = message.from_user.id
    settings = await get_settings(user_id)
    photo = message.photo[-1]
    tg_file = await bot.get_file(photo.file_id)
    buf = BytesIO()
    await bot.download_file(tg_file.file_path, destination=buf)
    caption = message.caption or ""
    await message.answer("🔍 Анализирую фото...")
    result = await analyze_image(buf.getvalue(), caption)
    await send_reply(message, result, settings)

@router.message(F.document)
async def on_document(message: Message) -> None:
    user_id = message.from_user.id
    settings = await get_settings(user_id)
    doc = message.document
    filename = doc.file_name or "file"
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    tg_file = await bot.get_file(doc.file_id)
    buf = BytesIO()
    await bot.download_file(tg_file.file_path, destination=buf)
    data = buf.getvalue()
    text = ""
    try:
        if ext == "pdf":
            reader = PdfReader(BytesIO(data))
            text = "\n".join(p.extract_text() or "" for p in reader.pages)
        elif ext == "docx":
            d = DocxDocument(BytesIO(data))
            text = "\n".join(p.text for p in d.paragraphs if p.text.strip())
        elif ext in {"xlsx", "xlsm"}:
            wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
            parts = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    cells = [str(c) for c in row if c is not None]
                    if cells:
                        parts.append(" | ".join(cells))
            text = "\n".join(parts)
        else:
            text = data.decode("utf-8", errors="ignore")
    except Exception as e:
        await message.answer(f"❌ Ошибка разбора документа: {e}")
        return

    text = text.strip()
    if not text:
        await message.answer("📂 Документ пуст или не содержит текста.")
        return

    await message.answer(f"📂 Документ «{filename}» загружен ({len(text)} симв.). Обрабатываю контекст...")
    prompt = f"Контекст из файла {filename}:\n\"\"\"\n{text[:4000]}\n\"\"\"\n\nПроанализируй данный документ."
    reply = await jarvis_reply(user_id, prompt, settings, save=False)
    await send_reply(message, reply, settings)

@router.message(F.text)
async def on_text(message: Message) -> None:
    user_id = message.from_user.id
    settings = await get_settings(user_id)
    text = message.text.strip()
    if text.startswith("/"):
        return
    if await handle_intents(message, text, settings):
        return
    reply = await jarvis_reply(user_id, text, settings)
    await send_reply(message, reply, settings)

# ══════════════════════════════════════════════════════════════
# HTTP / API И КОНТРОЛЬ ЗАДАЧ ИЗВНЕ
# ══════════════════════════════════════════════════════════════
async def http_jarvis(request: web.Request) -> web.Response:
    try:
        auth = request.headers.get("Authorization", "")
        if JARVIS_API_KEY and auth != f"Bearer {JARVIS_API_KEY}":
            return web.json_response({"error": "Unauthorized"}, status=401)
        data = await request.json()
        text = data.get("text", "").strip()
        user_id = int(data.get("user_id") or JARVIS_HTTP_USER_ID)
        if not text or not user_id:
            return web.json_response({"error": "Missing parameters"}, status=400)
        settings = await get_settings(user_id)
        answer = await jarvis_reply(user_id, text, settings)
        return web.json_response({"reply": answer})
    except Exception as e:
        return web.json_response({"error": str(e)}, status=500)

async def http_health(request: web.Request) -> web.Response:
    return web.json_response({"status": "ok", "time": now_str()})

# ══════════════════════════════════════════════════════════════
# ШЕДУЛЕРНЫЕ ЗАДАЧИ
# ══════════════════════════════════════════════════════════════
async def reminders_job() -> None:
    for r in await reminder_get_due():
        try:
            await bot.send_message(r["chat_id"], f"⏰ Напоминание:\n{r['text']}\n\n({r['due_at']})")
            await reminder_mark_sent(r["id"])
        except Exception as e:
            logger.warning("Reminder send error: %s", e)

async def briefing_job() -> None:
    now = datetime.now()
    today = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M")
    rows = await db_query("SELECT user_id, briefing_time, last_brief_date FROM settings WHERE briefing_enabled=1")
    for u in rows:
        try:
            if u["briefing_time"] != current_time:
                continue
            if u.get("last_brief_date") == today:
                continue
            settings = await get_settings(u["user_id"])
            city = settings.get("default_city") or "Ханчжоу"
            weather = await get_weather(city)
            tasks = await task_list(u["user_id"], limit=5)
            health = await health_summary(u["user_id"])
            lines = [f"☀️ Автоматический Брифинг", "", weather, ""]
            if tasks:
                lines.append(f"📋 Актуальные задачи:")
                for t in tasks[:3]:
                    lines.append(f" • #{t['id']} {t['text']}")
            lines.append(f"\n{health}")
            await bot.send_message(u["user_id"], "\n".join(lines))
            await db_execute("UPDATE settings SET last_brief_date=? WHERE user_id=?", (today, u["user_id"]))
        except Exception as e:
            logger.error("Briefing job error for user %s: %s", u["user_id"], e)

# ══════════════════════════════════════════════════════════════
# ЗАПУСК СЕРВЕРА И БОТА
# ══════════════════════════════════════════════════════════════
async def main() -> None:
    init_db()
    scheduler.add_job(reminders_job, "interval", seconds=30, id="reminders", replace_existing=True)
    scheduler.add_job(briefing_job, "interval", seconds=60, id="briefing", replace_existing=True)
    scheduler.start()

    port = int(os.environ.get("PORT", "8080"))
    app = web.Application()
    app.router.add_post("/jarvis", http_jarvis)
    app.router.add_get("/health", http_health)
    runner = web.AppRunner(app)
    await runner.setup()
    await web.TCPSite(runner, "0.0.0.0", port).start()
    logger.info(f"HTTP Server started on port {port}")

    try:
        await dp.start_polling(bot)
    finally:
        await bot.session.close()
        global HTTP_SESSION
        if HTTP_SESSION and not HTTP_SESSION.closed:
            await HTTP_SESSION.close()

if __name__ == "__main__":
    asyncio.run(main())
