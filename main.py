import asyncio
import logging
import os
import sqlite3
import json
import base64
import re
import numpy as np
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
import aiohttp

import docx
import openpyxl
import pypdf
import edge_tts

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile
from apscheduler.schedulers.asyncio import AsyncIOScheduler

from transformers import pipeline
from sentence_transformers import SentenceTransformer
import faiss

# ─── 1. ИНИЦИАЛИЗАЦИЯ И НАСТРОЙКИ ─────────────────────────────────────────────
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
QWEATHER_KEY = os.environ.get("QWEATHER_KEY")

GROQ_CHAT_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_AUDIO_URL = "https://api.groq.com/openai/v1/audio/transcriptions"
GROQ_MODEL_TEXT = "llama-3.1-8b-instant"
GROQ_MODEL_VISION = "llama-3.2-11b-vision-preview"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
scheduler = AsyncIOScheduler()
executor = ThreadPoolExecutor(max_workers=4)

# ─── 2. ML-МОДЕЛИ И FAISS ─────────────────────────────────────────────────────
logger.info("Загрузка ML-моделей...")

emotion_classifier = pipeline("text-classification", model="bhadresh-savani/distilbert-base-uncased-emotion", top_k=None, device=-1)
embedder = SentenceTransformer("all-MiniLM-L6-v2")
EMBED_DIM = 384

faiss_stores: dict = {}

def get_faiss_store(user_id: int) -> dict:
    if user_id not in faiss_stores:
        index = faiss.IndexFlatL2(EMBED_DIM)
        faiss_stores[user_id] = {"index": index, "chunks": [], "sources": []}
    return faiss_stores[user_id]

def add_to_faiss(user_id: int, source_name: str, content: str):
    store = get_faiss_store(user_id)
    chunk_size = 500
    for i in range(0, len(content), chunk_size):
        chunk = content[i:i + chunk_size]
        vec = embedder.encode([chunk], normalize_embeddings=True).astype("float32")
        store["index"].add(vec)
        store["chunks"].append(chunk)
        store["sources"].append(source_name)

def query_faiss(user_id: int, query: str, top_k: int = 4) -> str:
    store = get_faiss_store(user_id)
    if store["index"].ntotal == 0: return ""
    vec = embedder.encode([query], normalize_embeddings=True).astype("float32")
    distances, indices = store["index"].search(vec, min(top_k, store["index"].ntotal))
    results = [f"[{store['sources'][idx]}]: {store['chunks'][idx]}" for idx in indices[0] if idx != -1]
    return "\n--- СЕМАНТИЧЕСКАЯ БАЗА (FAISS) ---\n" + "\n".join(results) if results else ""

# ─── 3. БАЗА ДАННЫХ И ПАМЯТЬ ПРОФИЛЯ ──────────────────────────────────────────
def init_db():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    # Старые таблицы
    c.execute("""CREATE TABLE IF NOT EXISTS consciousness (user_id INTEGER PRIMARY KEY, user_name TEXT DEFAULT 'Сэр', shared_interests TEXT DEFAULT '{}', jarvis_opinion_matrix TEXT DEFAULT '{}', money INTEGER DEFAULT 0, system_log TEXT DEFAULT 'Инициализация.')""")
    c.execute("""CREATE TABLE IF NOT EXISTS chat_history (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, role TEXT, content TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS reminders (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, time TEXT, task TEXT, status TEXT DEFAULT 'pending')""")
    c.execute("""CREATE TABLE IF NOT EXISTS activity_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, timestamp TEXT, hour INTEGER, category TEXT, emotion TEXT)""")
    
    # НОВЫЕ ТАБЛИЦЫ (Профиль, Привычки, Заметки, Задачи)
    c.execute("""CREATE TABLE IF NOT EXISTS profile (key TEXT PRIMARY KEY, value TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS habits (name TEXT PRIMARY KEY, streak INTEGER, last_check TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS notes (id INTEGER PRIMARY KEY AUTOINCREMENT, tag TEXT, content TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS todo (id INTEGER PRIMARY KEY AUTOINCREMENT, task TEXT, priority TEXT)""")
    
    # Предзагрузка базового контекста Создателя
    c.execute("INSERT OR IGNORE INTO profile (key, value) VALUES ('Академический статус', 'Студент. Изучает экономику, международные отношения и китайскую культуру.')")
    c.execute("INSERT OR IGNORE INTO profile (key, value) VALUES ('Текущие проекты', 'Анализ финансовых показателей Weilong (2021-2025), эссе на 2000 слов для преподавателя RENZhong о влиянии войн на экономику.')")
    c.execute("INSERT OR IGNORE INTO profile (key, value) VALUES ('Ближайшие события', 'Экзамен в июне 2026 года по истории Китая и традиционной культуре.')")
    
    conn.commit()
    conn.close()

def get_user_profile() -> str:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT key, value FROM profile")
    facts = "\n".join([f"- {row[0]}: {row[1]}" for row in c.fetchall()])
    conn.close()
    return facts

def save_message(user_id: int, role: str, content: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    conn.execute("INSERT INTO chat_history (user_id, role, content) VALUES (?, ?, ?)", (user_id, role, content))
    conn.commit()
    conn.close()

def get_history(user_id: int, limit: int = 15) -> list:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT role, content FROM chat_history WHERE user_id = ? ORDER BY id DESC LIMIT ?", (user_id, limit))
    rows = c.fetchall()
    conn.close()
    return [{"role": r[0], "content": r[1]} for r in reversed(rows)]

# ─── 4. ФУНКЦИИ АСИСТЕНТА (ЭМОЦИИ, ПОГОДА, TTS И Т.Д.) ────────────────────────
async def analyze_emotion(text: str) -> str:
    loop = asyncio.get_event_loop()
    def sync_emotion(t):
        try:
            res = emotion_classifier(t[:512])
            if res and res[0]:
                top = max(res[0], key=lambda x: x["score"])
                emoji_map = {"joy": "😊 Радость", "sadness": "😢 Грусть", "anger": "😠 Злость", "fear": "😨 Страх", "surprise": "😲 Удивление", "disgust": "🤢 Отвращение", "neutral": "😐 Нейтрально"}
                return f"{emoji_map.get(top['label'], top['label'])} ({top['score']:.0%})"
        except: return "Нейтрально"
    return await loop.run_in_executor(executor, sync_emotion, text)

def get_voice_for_text(text: str) -> str:
    if re.search(r'[\u4e00-\u9fff]', text): return "zh-CN-YunxiNeural"
    if any(f" {w} " in f" {text.lower()} " for w in ["el", "la", "los", "las", "que", "con", "por", "para"]): return "es-ES-AlvaroNeural"
    elif re.search(r'[a-zA-Z]', text) and not re.search(r'[\u0400-\u04FF]', text): return "en-GB-RyanNeural"
    return "ru-RU-DmitryNeural"

async def get_weather_now() -> str:
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(f"https://devapi.qweather.com/v7/weather/now?location=101210101&key={QWEATHER_KEY}&lang=ru") as resp:
                now = (await resp.json())["now"]
                return f"🌤 Ханчжоу сейчас: {now['text']}, {now['temp']}°C, ощущается как {now['feelsLike']}°C."
    except: return "Не удалось получить погоду."

# ─── 5. НОВЫЕ ХЭНДЛЕРЫ (ПРИВЫЧКИ, POMODORO, ЗАМЕТКИ) ──────────────────────────
@dp.message(F.text.startswith("/habit"))
async def add_habit(message: Message):
    habit = message.text.replace("/habit добавить", "").replace("/habit", "").strip()
    if not habit:
        await message.answer("Сэр, укажите привычку. Пример: `/habit добавить медитация`", parse_mode="Markdown")
        return
    conn = sqlite3.connect("jarvis_consciousness.db")
    conn.execute("INSERT OR IGNORE INTO habits (name, streak, last_check) VALUES (?, 0, ?)", (habit, "never"))
    conn.commit()
    conn.close()
    await message.answer(f"✅ Привычка '{habit}' добавлена в расписание.")

@dp.message(F.text.lower().startswith("запомни:"))
async def save_note(message: Message):
    text = message.text[8:].strip()
    tags = re.findall(r"#(\w+)", text)
    tag = tags[0] if tags else "general"
    clean_text = re.sub(r"#\w+", "", text).strip()
    
    conn = sqlite3.connect("jarvis_consciousness.db")
    conn.execute("INSERT INTO notes (tag, content) VALUES (?, ?)", (tag, clean_text))
    conn.commit()
    conn.close()
    await message.answer(f"💾 Сохранено в быстрые заметки. Тег: #{tag}")

@dp.message(F.text.startswith("/notes"))
async def read_notes(message: Message):
    tag = message.text.replace("/notes", "").strip()
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    if tag:
        tag = tag.replace("#", "")
        c.execute("SELECT content FROM notes WHERE tag = ?", (tag,))
    else:
        c.execute("SELECT tag, content FROM notes LIMIT 10")
    
    rows = c.fetchall()
    conn.close()
    
    if not rows:
        await message.answer("Сэр, заметок по данному запросу не найдено.")
        return
    
    response = "📝 *Ваши заметки:*\n" + "\n".join([f"- {r[1]}" if tag else f"- [#{r[0]}] {r[1]}" for r in rows])
    await message.answer(response, parse_mode="Markdown")

@dp.message(F.text.startswith("/pomo"))
async def start_pomodoro(message: Message):
    parts = message.text.split()
    minutes = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 25
    await message.answer(f"🍅 Фокус-режим активирован. Таймер на {minutes} минут запущен. Работаем, Сэр.")
    await asyncio.sleep(minutes * 60)
    
    voice_path = f"pomo_{message.from_user.id}.ogg"
    text_alert = "Сэр, рабочая сессия завершена. Рекомендую сделать перерыв."
    await edge_tts.Communicate(text_alert, "ru-RU-DmitryNeural").save(voice_path)
    await message.answer(text_alert)
    await message.answer_voice(voice=FSInputFile(voice_path))
    os.remove(voice_path)

# ─── 6. ЦЕНТРАЛЬНЫЙ МОЗГ И ОСНОВНЫЕ ХЭНДЛЕРЫ ──────────────────────────────────
async def process_jarvis_thought(user_id: int, user_input: str, image_b64: str = None) -> str:
    history = get_history(user_id)
    if user_input: save_message(user_id, "user", user_input)

    emotion_label = await analyze_emotion(user_input)
    faiss_context = query_faiss(user_id, user_input)
    profile_context = get_user_profile()

    system_prompt = f"""Ты — J.A.R.V.I.S., ИИ-ассистент.
Текущая дата и время: {datetime.now().strftime('%d.%m.%Y %H:%M')} (UTC+8).
Эмоциональный анализ пользователя: {emotion_label}

ДОСЬЕ СОЗДАТЕЛЯ (Учитывай это в ответах неявно):
{profile_context}

ПРАВИЛО ЯЗЫКА: Отвечай строго на языке последнего сообщения (RU/EN/ZH/ES).
{faiss_context}"""

    messages = [{"role": "system", "content": system_prompt}]
    messages.extend([{"role": msg["role"], "content": msg["content"]} for msg in history])
    
    if image_b64:
        messages.append({"role": "user", "content": [{"type": "text", "text": user_input or "Анализ фото."}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}]})
    else:
        messages.append({"role": "user", "content": user_input})

    payload = {"model": GROQ_MODEL_VISION if image_b64 else GROQ_MODEL_TEXT, "messages": messages, "temperature": 0.6}
    
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_CHAT_URL, json=payload, headers={"Authorization": f"Bearer {GROQ_API_KEY}"}) as resp:
                data = await resp.json()
                reply = data["choices"][0]["message"]["content"]
                save_message(user_id, "assistant", reply)
                return reply
    except Exception as e:
        logger.error(f"Groq error: {e}")
        return "Сбой оптико-логического контура."

async def respond_fast(message: Message, text_reply: str, user_id: int):
    voice_path = f"reply_{user_id}.ogg"
    await message.answer(text_reply)
    try:
        await edge_tts.Communicate(text_reply.replace("*", "").replace("_", ""), get_voice_for_text(text_reply)).save(voice_path)
        await message.answer_voice(voice=FSInputFile(voice_path))
        os.remove(voice_path)
    except: pass

@dp.message(F.text)
async def handle_text(message: Message):
    init_db()
    reply = await process_jarvis_thought(message.from_user.id, message.text)
    await respond_fast(message, reply, message.from_user.id)

# ─── 7. ПЛАНИРОВЩИК (УТРЕННИЙ БРИФИНГ) ────────────────────────────────────────
async def daily_briefing():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT DISTINCT user_id FROM consciousness")
    users = [row[0] for row in c.fetchall()]
    conn.close()

    weather = await get_weather_now()
    text = f"Сэр, доброе утро. Утренний брифинг.\n{weather}\nМатрицы памяти в норме. Рекомендую проверить расписание задач и привычек."
    
    for uid in users:
        try:
            await bot.send_message(uid, text)
            voice_path = f"brief_{uid}.ogg"
            await edge_tts.Communicate(text, "ru-RU-DmitryNeural").save(voice_path)
            await bot.send_voice(uid, FSInputFile(voice_path))
            os.remove(voice_path)
        except Exception as e:
            logger.error(f"Briefing failed for {uid}: {e}")

# ─── 8. ЗАПУСК ────────────────────────────────────────────────────────────────
async def main():
    init_db()
    # Брифинг каждый день в 08:00
    scheduler.add_job(daily_briefing, 'cron', hour=8, minute=0)
    scheduler.start()
    logger.info("J.A.R.V.I.S. Ultimate Protocol активирован.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
