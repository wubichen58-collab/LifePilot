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
GROQ_MODEL_TEXT = "llama-3.3-70b-versatile" # Лучше подходит для Orchestrator
GROQ_MODEL_FAST = "llama-3.1-8b-instant"
GROQ_MODEL_VISION = "llama-3.2-11b-vision-preview"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
scheduler = AsyncIOScheduler(timezone="Asia/Tokyo")
executor = ThreadPoolExecutor(max_workers=4)

# ─── 2. ML-МОДЕЛИ И FAISS (ФАЙЛОВАЯ ПАМЯТЬ) ───────────────────────────────────
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
    
    # Логируем загрузку файла для утреннего дайджеста
    conn = sqlite3.connect("jarvis_consciousness.db")
    conn.execute("INSERT INTO indexed_files (user_id, filename, timestamp) VALUES (?, ?, ?)", 
                 (user_id, source_name, datetime.now().isoformat()))
    conn.commit()
    conn.close()

    for i in range(0, len(content), chunk_size):
        chunk = content[i:i + chunk_size]
        vec = embedder.encode([chunk], normalize_embeddings=True).astype("float32")
        store["index"].add(vec)
        store["chunks"].append(chunk)
        store["sources"].append(source_name)

async def query_faiss_async(user_id: int, query: str, top_k: int = 4) -> str:
    def _query():
        store = get_faiss_store(user_id)
        if store["index"].ntotal == 0: return ""
        vec = embedder.encode([query], normalize_embeddings=True).astype("float32")
        distances, indices = store["index"].search(vec, min(top_k, store["index"].ntotal))
        results = [f"[{store['sources'][idx]}]: {store['chunks'][idx]}" for idx in indices[0] if idx != -1]
        return "\n--- ЛОКАЛЬНЫЕ ФАЙЛЫ (FAISS) ---\n" + "\n".join(results) if results else ""
    
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(executor, _query)

# ─── 3. БАЗА ДАННЫХ И ПАМЯТЬ ПРОФИЛЯ ──────────────────────────────────────────
def init_db():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS consciousness (user_id INTEGER PRIMARY KEY, user_name TEXT DEFAULT 'Сэр', shared_interests TEXT DEFAULT '{}', jarvis_opinion_matrix TEXT DEFAULT '{}', money INTEGER DEFAULT 0, system_log TEXT DEFAULT 'Инициализация.')""")
    c.execute("""CREATE TABLE IF NOT EXISTS chat_history (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, role TEXT, content TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS reminders (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, time TEXT, task TEXT, status TEXT DEFAULT 'pending')""")
    c.execute("""CREATE TABLE IF NOT EXISTS activity_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, timestamp TEXT, hour INTEGER, category TEXT, emotion TEXT)""")
    
    # Расширенные таблицы
    c.execute("""CREATE TABLE IF NOT EXISTS profile (key TEXT PRIMARY KEY, value TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS habits (name TEXT PRIMARY KEY, streak INTEGER, last_check TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS notes (id INTEGER PRIMARY KEY AUTOINCREMENT, tag TEXT, content TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS todo (id INTEGER PRIMARY KEY AUTOINCREMENT, task TEXT, priority TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS indexed_files (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, filename TEXT, timestamp TEXT)""")
    
    # Вшитый актуальный контекст
    c.execute("INSERT OR IGNORE INTO profile (key, value) VALUES ('Академический профиль', 'Студент. Экономика, международные отношения, культура Китая, логистика (Новый Шелковый путь).')")
    c.execute("INSERT OR IGNORE INTO profile (key, value) VALUES ('Проект: Weilong', 'Анализ финансовых показателей Weilong (2021-2025).')")
    c.execute("INSERT OR IGNORE INTO profile (key, value) VALUES ('Проект: Эссе RENZhong', '2000 слов. Тема: Влияние войн между США, Израилем и Ираном на национальную экономику.')")
    c.execute("INSERT OR IGNORE INTO profile (key, value) VALUES ('События', 'Подготовка к экзамену в июне 2026 по истории Китая и культурному наследию.')")
    
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

# ─── 4. ФУНКЦИИ АСИСТЕНТА ─────────────────────────────────────────────────────
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
    if any(f" {w} " in f" {text.lower()} " for w in ["el", "la", "los", "las", "que"]): return "es-ES-AlvaroNeural"
    elif re.search(r'[a-zA-Z]', text) and not re.search(r'[\u0400-\u04FF]', text): return "en-GB-RyanNeural"
    return "ru-RU-DmitryNeural"

async def get_weather_now() -> str:
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(f"https://devapi.qweather.com/v7/weather/now?location=101210101&key={QWEATHER_KEY}&lang=ru") as resp:
                now = (await resp.json())["now"]
                return f"🌤 Погода: {now['text']}, {now['temp']}°C, ощущается как {now['feelsLike']}°C."
    except: return "Не удалось получить погоду."

async def call_groq_api(messages: list, model: str = GROQ_MODEL_TEXT, json_mode: bool = False) -> str:
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": 0.3}
    if json_mode: payload["response_format"] = {"type": "json_object"}
    
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            return data["choices"][0]["message"]["content"]

# ─── 5. ORCHESTRATOR И ЦЕНТРАЛЬНЫЙ МОЗГ ───────────────────────────────────────
async def process_jarvis_thought(user_id: int, user_input: str, image_b64: str = None) -> str:
    history = get_history(user_id)
    if user_input: save_message(user_id, "user", user_input)
    
    # Шаг 1: Агент Orchestrator (оценка сложности)
    orchestrator_prompt = f"""Ты — системный архитектор. Проанализируй запрос: "{user_input}".
    Требуется ли для этого сложный пошаговый план (например, написание эссе, анализ данных, сложный график)?
    Верни строго JSON: {{"plan_needed": true, "steps": ["Шаг 1: ...", "Шаг 2: ..."]}} или {{"plan_needed": false}}.
    """
    
    plan_text = ""
    try:
        plan_resp = await call_groq_api([{"role": "system", "content": orchestrator_prompt}], model=GROQ_MODEL_FAST, json_mode=True)
        plan_data = json.loads(plan_resp)
        if plan_data.get("plan_needed"):
            steps_joined = "\n".join(plan_data.get("steps", []))
            plan_text = f"\n\n[ВНУТРЕННЯЯ ДИРЕКТИВА ORCHESTRATOR]: Выполняй задачу строго по шагам:\n{steps_joined}"
            logger.info(f"Orchestrator activated: {steps_joined}")
    except Exception as e:
        logger.error(f"Orchestrator error: {e}")

    # Шаг 2: Сбор контекста (Эмоции + Локальные файлы)
    emotion_task = asyncio.create_task(analyze_emotion(user_input))
    faiss_task = asyncio.create_task(query_faiss_async(user_id, user_input))
    emotion_label, faiss_context = await asyncio.gather(emotion_task, faiss_task)
    
    # Шаг 3: Генерация финального ответа
    system_prompt = f"""Ты — J.A.R.V.I.S., ИИ-ассистент.
    Текущая дата: {datetime.now().strftime('%Y-%m-%d %H:%M')}.
    Эмоциональный фон: {emotion_label}.
    
    ДОСЬЕ СОЗДАТЕЛЯ:
    {get_user_profile()}
    {faiss_context}
    {plan_text}
    
    Отвечай лаконично, в стиле ИИ-дворецкого. При создании таблиц используй [GENERATE_DOC_XLSX], при отчетах [GENERATE_DOC_DOCX]."""

    messages = [{"role": "system", "content": system_prompt}]
    for msg in history: messages.append({"role": msg["role"], "content": msg["content"]})
    
    if image_b64:
        messages.append({"role": "user", "content": [{"type": "text", "text": user_input}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}]})
        model_to_use = GROQ_MODEL_VISION
    else:
        messages.append({"role": "user", "content": user_input})
        model_to_use = GROQ_MODEL_TEXT

    reply = await call_groq_api(messages, model=model_to_use)
    save_message(user_id, "assistant", reply)
    return reply

# ─── 6. ХЭНДЛЕРЫ ДОКУМЕНТОВ И СООБЩЕНИЙ ──────────────────────────────────────
@dp.message(F.document)
async def handle_document(message: Message):
    user_id = message.from_user.id
    init_db()
    file_name = message.document.file_name
    file_ext = file_name.split(".")[-1].lower() if "." in file_name else ""
    local_tmp_path = f"tmp_{user_id}_{file_name}"
    
    await bot.download_file((await bot.get_file(message.document.file_id)).file_path, local_tmp_path)
    content = ""
    await message.answer(f"⏳ Сканирую документ `{file_name}`...")
    
    try:
        if file_ext in ["txt", "csv", "json"]:
            with open(local_tmp_path, "r", encoding="utf-8", errors="ignore") as f: content = f.read()
        elif file_ext == "docx":
            doc = docx.Document(local_tmp_path)
            content = "\n".join([p.text for p in doc.paragraphs])
        elif file_ext == "pdf":
            reader = pypdf.PdfReader(local_tmp_path)
            content = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
        elif file_ext in ["xlsx", "xls"]:
            wb = openpyxl.load_workbook(local_tmp_path, data_only=True)
            content = "\n".join([f"Лист {s.title}: " + " | ".join([str(c) for c in r if c]) for s in wb.worksheets for r in s.iter_rows(values_only=True)])

        if content:
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(executor, add_to_faiss, user_id, file_name, content)
            await message.answer(f"✅ База данных обновлена. Содержимое `{file_name}` загружено в векторную память (FAISS).")
        else:
            await message.answer("⚠️ Текст не обнаружен.")
    except Exception as e:
        logger.error(f"Doc error: {e}")
        await message.answer(f"💥 Ошибка чтения файла.")
    finally:
        if os.path.exists(local_tmp_path): os.remove(local_tmp_path)

@dp.message(F.text.startswith("/notes"))
async def read_notes(message: Message):
    tag = message.text.replace("/notes", "").strip()
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    if tag:
        c.execute("SELECT content FROM notes WHERE tag = ?", (tag.replace("#", ""),))
    else:
        c.execute("SELECT tag, content FROM notes LIMIT 10")
    rows = c.fetchall()
    conn.close()
    
    if not rows: return await message.answer("Сэр, заметок нет.")
    response = "📝 *Ваши заметки:*\n" + "\n".join([f"- {r[0]}" if tag else f"- [#{r[0]}] {r[1]}" for r in rows])
    await message.answer(response, parse_mode="Markdown")

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    init_db()
    
    if message.text.lower().startswith("запомни:"):
        text = message.text[8:].strip()
        tags = re.findall(r"#(\w+)", text)
        tag = tags[0] if tags else "general"
        clean_text = re.sub(r"#\w+", "", text).strip()
        conn = sqlite3.connect("jarvis_consciousness.db")
        conn.execute("INSERT INTO notes (tag, content) VALUES (?, ?)", (tag, clean_text))
        conn.commit()
        return await message.answer(f"💾 Зафиксировано в памяти. Тег: #{tag}")

    reply = await process_jarvis_thought(user_id, message.text)
    
    # Генерация TTS
    voice_path = f"reply_{user_id}.ogg"
    try:
        dynamic_voice = get_voice_for_text(reply)
        clean_reply = re.sub(r'\[GENERATE_DOC_.*?\]', '', reply)
        await edge_tts.Communicate(clean_reply.replace("*", "").replace("_", ""), dynamic_voice).save(voice_path)
        await message.answer(clean_reply)
        if os.path.exists(voice_path):
            await message.answer_voice(voice=FSInputFile(voice_path))
            os.remove(voice_path)
    except Exception:
        await message.answer(reply)

# ─── 7. CRON ДАЙДЖЕСТ (БРИФИНГ) ───────────────────────────────────────────────
async def daily_briefing():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT DISTINCT user_id FROM consciousness")
    users = [row[0] for row in c.fetchall()]
    
    # Проверка проиндексированных файлов за последние 24 часа
    yesterday = (datetime.now() - timedelta(days=1)).isoformat()
    c.execute("SELECT filename FROM indexed_files WHERE timestamp > ?", (yesterday,))
    recent_files = [r[0] for r in c.fetchall()]
    conn.close()

    weather = await get_weather_now()
    files_str = f"Проиндексировано новых файлов: {len(recent_files)} ({', '.join(recent_files[:3])})." if recent_files else "Новых файлов в базе нет."
    
    text = f"Сэр, доброе утро. Автоматический дайджест систем.\n{weather}\n{files_str}\nМатрицы FAISS активны. Жду ваших указаний по текущим проектам."
    
    for uid in users:
        try:
            await bot.send_message(uid, text)
            voice_path = f"brief_{uid}.ogg"
            await edge_tts.Communicate(text, "ru-RU-DmitryNeural").save(voice_path)
            await bot.send_voice(uid, FSInputFile(voice_path))
            os.remove(voice_path)
        except Exception as e: logger.error(f"Briefing failed: {e}")

# ─── 8. ЗАПУСК ────────────────────────────────────────────────────────────────
async def main():
    init_db()
    # Утренний брифинг в 08:00
    scheduler.add_job(daily_briefing, 'cron', hour=8, minute=0)
    scheduler.start()
    logger.info("J.A.R.V.I.S. Ultimate с Orchestrator активен.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
