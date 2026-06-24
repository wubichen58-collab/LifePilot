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

from aiogram import Bot, Dispatcher, types, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import Command
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

# ─── 2. ML-МОДЕЛИ ─────────────────────────────────────────────────────────────
logger.info("Загрузка ML-моделей...")

emotion_classifier = pipeline(
    "text-classification",
    model="bhadresh-savani/distilbert-base-uncased-emotion",
    top_k=None,
    device=-1
)

embedder = SentenceTransformer("all-MiniLM-L6-v2")
EMBED_DIM = 384

logger.info("ML-модели загружены.")

# ─── 3. ФЕЙС-КОНТРОЛЬ (заглушка — DeepFace несовместим с Railway) ─────────────
def verify_face_sync(img_path: str) -> bool:
    return True

async def verify_face(img_path: str) -> bool:
    return True

# ─── 4. FAISS — ВЕКТОРНАЯ БАЗА ЗНАНИЙ ─────────────────────────────────────────
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
    if store["index"].ntotal == 0:
        return ""
    vec = embedder.encode([query], normalize_embeddings=True).astype("float32")
    distances, indices = store["index"].search(vec, min(top_k, store["index"].ntotal))
    results = []
    for idx in indices[0]:
        if idx != -1:
            results.append(f"[{store['sources'][idx]}]: {store['chunks'][idx]}")
    return "\n--- СЕМАНТИЧЕСКАЯ БАЗА (FAISS) ---\n" + "\n".join(results) if results else ""

# ─── 5. АНАЛИЗ ЭМОЦИЙ ─────────────────────────────────────────────────────────
def analyze_emotion_sync(text: str) -> str:
    try:
        results = emotion_classifier(text[:512])
        if results and results[0]:
            top = max(results[0], key=lambda x: x["score"])
            emoji_map = {
                "joy": "😊 Радость", "sadness": "😢 Грусть", "anger": "😠 Злость",
                "fear": "😨 Страх", "surprise": "😲 Удивление",
                "disgust": "🤢 Отвращение", "neutral": "😐 Нейтрально"
            }
            return f"{emoji_map.get(top['label'], top['label'])} ({top['score']:.0%})"
    except Exception as e:
        logger.error(f"Emotion error: {e}")
    return "Нейтрально"

async def analyze_emotion(text: str) -> str:
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(executor, analyze_emotion_sync, text)

# ─── 6. КИТАЙСКИЕ СЕРВИСЫ ─────────────────────────────────────────────────────
async def get_weather_now() -> str:
    try:
        async with aiohttp.ClientSession() as session:
            url = f"https://devapi.qweather.com/v7/weather/now?location=101210101&key={QWEATHER_KEY}&lang=ru"
            async with session.get(url, timeout=aiohttp.ClientTimeout(total=5)) as resp:
                data = await resp.json()
                now = data["now"]
                return (f"🌤 Ханчжоу сейчас: {now['text']}, {now['temp']}°C, "
                        f"ощущается как {now['feelsLike']}°C, "
                        f"влажность {now['humidity']}%, ветер {now['windSpeed']} км/ч")
    except Exception as e:
        logger.error(f"QWeather error: {e}")
        return "Не удалось получить погоду."

async def get_weather_forecast() -> str:
    try:
        async with aiohttp.ClientSession() as session:
            url = f"https://devapi.qweather.com/v7/weather/3d?location=101210101&key={QWEATHER_KEY}&lang=ru"
            async with session.get(url, timeout=aiohttp.ClientTimeout(total=5)) as resp:
                data = await resp.json()
                lines = ["📅 Прогноз на 3 дня:"]
                for day in data["daily"]:
                    lines.append(f"{day['fxDate']}: {day['textDay']}, {day['tempMin']}–{day['tempMax']}°C")
                return "\n".join(lines)
    except Exception as e:
        logger.error(f"QWeather forecast error: {e}")
        return "Не удалось получить прогноз."

async def search_netease_music(query: str) -> str:
    try:
        async with aiohttp.ClientSession() as session:
            url = "https://music.163.com/api/search/get"
            params = {"s": query, "type": 1, "limit": 5}
            headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://music.163.com"}
            async with session.post(url, data=params, headers=headers,
                                    timeout=aiohttp.ClientTimeout(total=5)) as resp:
                data = await resp.json(content_type=None)
                songs = data["result"]["songs"]
                lines = [f"🎵 Результаты по '{query}':"]
                for s in songs:
                    artist = s["artists"][0]["name"]
                    lines.append(f"• {s['name']} — {artist}")
                return "\n".join(lines)
    except Exception as e:
        logger.error(f"NetEase error: {e}")
        return "Не удалось найти музыку."

# ─── 7. БАЗА ДАННЫХ ───────────────────────────────────────────────────────────
def init_db():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS consciousness (
        user_id INTEGER PRIMARY KEY, user_name TEXT DEFAULT 'Сэр',
        shared_interests TEXT DEFAULT '{}', jarvis_opinion_matrix TEXT DEFAULT '{}',
        money INTEGER DEFAULT 0, system_log TEXT DEFAULT 'Инициализация.')""")
    c.execute("""CREATE TABLE IF NOT EXISTS chat_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, role TEXT, content TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS reminders (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER,
        time TEXT, task TEXT, status TEXT DEFAULT 'pending')""")
    c.execute("""CREATE TABLE IF NOT EXISTS activity_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER,
        timestamp TEXT, hour INTEGER, category TEXT, emotion TEXT)""")
    conn.commit()
    conn.close()

def get_mind(user_id: int) -> dict:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT user_id, user_name, system_log FROM consciousness WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    if not row:
        c.execute("INSERT OR IGNORE INTO consciousness (user_id) VALUES (?)", (user_id,))
        conn.commit()
        c.execute("SELECT user_id, user_name, system_log FROM consciousness WHERE user_id = ?", (user_id,))
        row = c.fetchone()
    conn.close()
    return {"user_id": row[0], "user_name": row[1], "system_log": row[2]}

def save_message(user_id: int, role: str, content: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("INSERT INTO chat_history (user_id, role, content) VALUES (?, ?, ?)", (user_id, role, content))
    conn.commit()
    conn.close()

def get_history(user_id: int, limit: int = 20) -> list:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT role, content FROM chat_history WHERE user_id = ? ORDER BY id DESC LIMIT ?", (user_id, limit))
    rows = c.fetchall()
    conn.close()
    return [{"role": r[0], "content": r[1]} for r in reversed(rows)]

def log_user_activity(user_id: int, category: str, emotion: str = ""):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("INSERT INTO activity_logs (user_id, timestamp, hour, category, emotion) VALUES (?, ?, ?, ?, ?)",
              (user_id, datetime.now().isoformat(), datetime.now().hour, category, emotion))
    conn.commit()
    conn.close()

def add_reminder_db(user_id: int, run_time: datetime, task: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("INSERT INTO reminders (user_id, time, task) VALUES (?, ?, ?)", (user_id, run_time.isoformat(), task))
    conn.commit()
    conn.close()

# ─── 8. WEB ПОИСК И ДОКУМЕНТЫ ─────────────────────────────────────────────────
async def search_web_ddg(query: str) -> str:
    url = f"https://html.duckduckgo.com/html/?q={query}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers, timeout=aiohttp.ClientTimeout(total=5)) as resp:
                if resp.status != 200:
                    return ""
                html = await resp.text()
                snippets = re.findall(r'<a class="result__snippet".*?>(.*?)</a>', html, re.DOTALL)
                clean_snips = [re.sub(r'<[^>]*>', '', s).strip() for s in snippets[:3]]
                return "\n--- ИНТЕРНЕТ ---\n" + "\n".join(clean_snips) if clean_snips else ""
    except:
        return ""

def create_docx_report(title: str, content: str, filename: str):
    doc = docx.Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"Дата генерации: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\nСистема: J.A.R.V.I.S.\n")
    doc.add_paragraph(content)
    doc.save(filename)

def create_xlsx_report(table_data: list, filename: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in table_data:
        ws.append(row)
    wb.save(filename)

def escape_markdown(text: str) -> str:
    return re.sub(r'([_*\[\]()~`>#+\-=|{}.!])', r'\\\1', text)

def get_voice_for_text(text: str) -> str:
    if re.search(r'[\u4e00-\u9fff]', text):
        return "zh-CN-YunxiNeural"
    elif re.search(r'[a-zA-Z]', text) and not re.search(r'[\u0400-\u04FF]', text):
        return "en-GB-RyanNeural"
    return "ru-RU-DmitryNeural"

# ─── 9. ПЛАНИРОВЩИК ЗАДАЧ ─────────────────────────────────────────────────────
async def send_reminder(user_id: int, task: str):
    try:
        text = f"🔔 *Напоминание:* Сэр, {task}"
        await bot.send_message(chat_id=user_id, text=escape_markdown(text), parse_mode="MarkdownV2")
        voice_path = f"remind_{user_id}.ogg"
        dynamic_voice = get_voice_for_text(task)
        await edge_tts.Communicate(f"Сэр, напоминаю: {task}", dynamic_voice).save(voice_path)
        if os.path.exists(voice_path):
            await bot.send_voice(chat_id=user_id, voice=FSInputFile(voice_path))
            os.remove(voice_path)
    except Exception as e:
        logger.error(f"Scheduler error: {e}")

async def check_and_extract_reminders(user_id: int, user_input: str):
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    sys_inst = 'Если в тексте есть задача на будущее, верни JSON: {"has_reminder": true, "minutes_delay": число, "task": "суть"}. Иначе: {"has_reminder": false}'
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [{"role": "system", "content": sys_inst}, {"role": "user", "content": user_input}],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
                data = await resp.json()
                res = json.loads(data["choices"][0]["message"]["content"])
                if res.get("has_reminder"):
                    delay = int(res.get("minutes_delay", 0))
                    task = res.get("task", "Важное дело")
                    if delay > 0:
                        run_time = datetime.now() + timedelta(minutes=delay)
                        add_reminder_db(user_id, run_time, task)
                        scheduler.add_job(send_reminder, "date", run_date=run_time, args=[user_id, task])
                        return f"\n\n⚡️ *[Память обновлена: «{task}» через {delay} мин.]*"
    except:
        pass
    return ""

# ─── 10. ЦЕНТРАЛЬНЫЙ МОЗГ ─────────────────────────────────────────────────────
async def process_jarvis_thought(user_id: int, user_input: str, image_b64: str = None) -> str:
    init_db()
    history = get_history(user_id, limit=20)

    if user_input:
        save_message(user_id, "user", user_input)

    emotion_task = asyncio.create_task(analyze_emotion(user_input))
    reminder_task = asyncio.create_task(check_and_extract_reminders(user_id, user_input))

    # Определяем контекст из китайских сервисов или веба
    web_context = ""
    lower = user_input.lower()
    if any(w in lower for w in ["погода", "weather", "天气", "дождь", "температура"]):
        web_context = await get_weather_now()
    elif any(w in lower for w in ["прогноз", "forecast", "预报", "завтра"]):
        web_context = await get_weather_forecast()
    elif any(w in lower for w in ["музыка", "песня", "music", "歌曲", "найди песню", "网易"]):
        # Извлекаем запрос после ключевого слова
        query = re.sub(r'.*(музыка|песня|music|найди песню)\s*', '', user_input, flags=re.IGNORECASE).strip()
        if not query:
            query = user_input
        web_context = await search_netease_music(query)
    elif any(w in lower for w in ["новости", "найти", "сейчас", "курс", "интернет"]):
        web_context = await search_web_ddg(user_input)

    emotion_label = await emotion_task
    log_user_activity(user_id, "chat", emotion_label)
    faiss_context = query_faiss(user_id, user_input)

    system_prompt = f"""Ты — J.A.R.V.I.S., ИИ-ассистент Создателя.
Текущая дата и время: {datetime.now().strftime('%d.%m.%Y %H:%M')} (UTC+8, Ханчжоу).
Эмоциональный анализ: {emotion_label}
ПРАВИЛО ЯЗЫКА: Всегда отвечай на том языке, на котором написано последнее сообщение (RU/EN/ZH). Не смешивай без просьбы.
Документы: используй [GENERATE_DOC_DOCX] или [GENERATE_DOC_XLSX] если просят отчёт/таблицу.
{faiss_context}
{web_context}"""

    messages = [{"role": "system", "content": system_prompt}]
    for msg in history:
        messages.append({"role": msg["role"], "content": msg["content"]})

    if image_b64:
        messages.append({
            "role": "user",
            "content": [
                {"type": "text", "text": user_input or "Опиши подробно, что изображено на фото."},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}
            ]
        })
    else:
        messages.append({"role": "user", "content": user_input})

    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": GROQ_MODEL_VISION if image_b64 else GROQ_MODEL_TEXT,
        "messages": messages,
        "temperature": 0.6
    }

    async with aiohttp.ClientSession() as session:
        try:
            async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
                data = await resp.json()
                if "choices" in data:
                    reply = data["choices"][0]["message"]["content"]
                    save_message(user_id, "assistant", reply)
                    return reply + await reminder_task
                else:
                    logger.error(f"API Error: {data}")
                    return "Сбой оптико-логического контура."
        except Exception as e:
            logger.error(f"Groq API connection error: {e}")
            return "Сэр, нет связи с серверами."

async def respond_fast(message: Message, text_reply: str, user_id: int, custom_text_log: str = None):
    voice_path = f"reply_{user_id}.ogg"
    display_text = custom_text_log if custom_text_log else text_reply

    if "[GENERATE_DOC_DOCX]" in text_reply:
        clean_content = text_reply.replace("[GENERATE_DOC_DOCX]", "").strip()
        if len(clean_content) > 10:
            fn = f"report_{user_id}.docx"
            create_docx_report("Документ J.A.R.V.I.S.", clean_content, fn)
            await message.answer_document(FSInputFile(fn), caption="Сэр, ваш документ Word.")
            if os.path.exists(fn):
                os.remove(fn)
        display_text = display_text.replace("[GENERATE_DOC_DOCX]", "")
        text_reply = text_reply.replace("[GENERATE_DOC_DOCX]", "")

    elif "[GENERATE_DOC_XLSX]" in text_reply:
        lines = text_reply.replace("[GENERATE_DOC_XLSX]", "").strip().split("\n")
        table = [line.split("|") for line in lines if line]
        if len(table) > 0:
            fn = f"matrix_{user_id}.xlsx"
            create_xlsx_report(table, fn)
            await message.answer_document(FSInputFile(fn), caption="Сэр, таблица Excel.")
            if os.path.exists(fn):
                os.remove(fn)
        display_text = display_text.replace("[GENERATE_DOC_XLSX]", "")
        text_reply = text_reply.replace("[GENERATE_DOC_XLSX]", "")

    try:
        await message.answer(escape_markdown(display_text), parse_mode="MarkdownV2")
    except Exception:
        await message.answer(display_text, parse_mode=None)

    async def generate_and_send_voice():
        try:
            dynamic_voice = get_voice_for_text(text_reply)
            await edge_tts.Communicate(text_reply.replace("*", "").replace("_", ""), dynamic_voice).save(voice_path)
            if os.path.exists(voice_path) and os.path.getsize(voice_path) > 0:
                await message.answer_voice(voice=FSInputFile(voice_path))
                os.remove(voice_path)
        except Exception as e:
            logger.error(f"Voice generation failed: {e}")

    asyncio.create_task(generate_and_send_voice())

async def process_and_transcribe_audio(file_id: int, user_id: int) -> str:
    local_path = f"audio_{user_id}.ogg"
    try:
        file = await bot.get_file(file_id)
        await bot.download_file(file.file_path, local_path)
        headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
        data = aiohttp.FormData()
        data.add_field("file", open(local_path, "rb"), filename="voice.ogg")
        data.add_field("model", "whisper-large-v3")
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_AUDIO_URL, data=data, headers=headers) as resp:
                res = await resp.json()
                return res.get("text", "")
    except Exception as e:
        logger.error(f"Audio processing error: {e}")
        return ""
    finally:
        if os.path.exists(local_path):
            os.remove(local_path)

# ─── 11. ОБРАБОТЧИКИ ──────────────────────────────────────────────────────────
@dp.message(F.photo)
async def handle_photo(message: Message):
    user_id = message.from_user.id
    init_db()
    photo = message.photo[-1]
    local_path = f"img_{user_id}.jpg"
    await bot.download(photo, destination=local_path)

    face_verified = await verify_face(local_path)
    face_status = "✅ Личность подтверждена." if face_verified else "⚠️ Неизвестное лицо."
    await message.answer(f"📸 Фото получено. {face_status}\nВключаю визуальный анализ...")

    with open(local_path, "rb") as img_file:
        img_b64 = base64.b64encode(img_file.read()).decode("utf-8")
    os.remove(local_path)

    caption = message.caption or "Опиши подробно, что изображено на фото."
    reply = await process_jarvis_thought(user_id, caption, image_b64=img_b64)
    await respond_fast(message, reply, user_id)


@dp.message(F.voice | F.audio)
async def handle_audio(message: Message):
    user_id = message.from_user.id
    init_db()
    await message.answer("🎧 Распознаю аудиопоток...")
    file_id = message.voice.file_id if message.voice else message.audio.file_id
    text_input = await process_and_transcribe_audio(file_id, user_id)
    if text_input:
        reply = await process_jarvis_thought(user_id, f"[Аудио]: {text_input}")
        await respond_fast(message, reply, user_id, custom_text_log=f"🗣 *Распознано:* {text_input}\n\n{reply}")
    else:
        await message.answer("❌ Не удалось расшифровать аудио.")


@dp.message(F.document)
async def handle_document(message: Message):
    user_id = message.from_user.id
    init_db()
    file_name = message.document.file_name
    file_ext = file_name.split(".")[-1].lower() if "." in file_name else ""
    local_tmp_path = f"tmp_{user_id}_{file_name}"
    await bot.download_file((await bot.get_file(message.document.file_id)).file_path, local_tmp_path)
    content = ""
    try:
        if file_ext in ["txt", "csv", "json"]:
            with open(local_tmp_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
        elif file_ext == "docx":
            doc = docx.Document(local_tmp_path)
            content = "\n".join([p.text for p in doc.paragraphs])
        elif file_ext == "pdf":
            reader = pypdf.PdfReader(local_tmp_path)
            content = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
        elif file_ext in ["xlsx", "xls"]:
            wb = openpyxl.load_workbook(local_tmp_path, data_only=True)
            content = "\n".join([
                f"Лист {s.title}: " + " | ".join([str(c) for c in r if c])
                for s in wb.worksheets for r in s.iter_rows(values_only=True)
            ])
        else:
            with open(local_tmp_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read(5000)

        if content:
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(executor, add_to_faiss, user_id, file_name, content)
            await message.answer(f"✅ Файл `{file_name}` загружен в семантическую базу (FAISS).")
        else:
            await message.answer("⚠️ Не удалось извлечь текст из файла.")
    except Exception as e:
        logger.error(f"Document parse error: {e}")
        await message.answer(f"💥 Ошибка при анализе файла `{file_name}`.")
    finally:
        if os.path.exists(local_tmp_path):
            os.remove(local_tmp_path)


@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    init_db()
    reply = await process_jarvis_thought(user_id, message.text)
    await respond_fast(message, reply, user_id)


# ─── 12. ЗАПУСК ───────────────────────────────────────────────────────────────
async def main():
    init_db()
    scheduler.start()
    logger.info("J.A.R.V.I.S. с FAISS + Эмоции + QWeather + NetEase активен.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
