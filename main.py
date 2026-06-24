import asyncio
import logging
import os
import sqlite3
import json
import base64
import re
from datetime import datetime, timedelta
import aiohttp

from aiogram import Bot, Dispatcher, types, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import Command
import edge_tts

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from textblob import TextBlob
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.linear_model import LogisticRegression

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")

GROQ_CHAT_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_AUDIO_URL = "https://api.groq.com/openai/v1/audio/transcriptions"

GROQ_MODEL_TEXT = "llama-3.3-70b-versatile"
GROQ_MODEL_VISION = "llama-3.2-11b-vision-preview"

JARVIS_VOICE = "ru-RU-DmitryNeural" 

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
scheduler = AsyncIOScheduler()

# ─── МИНИ-ДВИЖОК МАШИННОГО ОБУЧЕНИЯ (ML) ────────────────────────────────────

TRAIN_TEXTS = [
    "надо срочно сделать отчет", "дедлайн завтра", "исправь ошибку в коде",
    "привет как дела", "расскажи шутку", "что делаешь",
    "я устал", "все надоело", "какой-то бред", "ничего не получается"
]
TRAIN_LABELS = [0, 0, 0, 1, 1, 1, 2, 2, 2, 2] # 0 - Work, 1 - Casual, 2 - Stress

vectorizer = CountVectorizer()
X_train = vectorizer.fit_transform(TRAIN_TEXTS)
ml_classifier = LogisticRegression()
ml_classifier.fit(X_train, TRAIN_LABELS)

def predict_context_ml(text: str) -> str:
    """Локально классифицирует тип запроса с помощью Machine Learning"""
    try:
        X_test = vectorizer.transform([text.lower()])
        prediction = ml_classifier.predict(X_test)[0]
        if prediction == 0: return "Деловой/Критический"
        if prediction == 2: return "Повышенный стресс/Усталость"
        return "Стандартный/Развлекательный"
    except:
        return "Не определен"

# ─── DATABASE SYSTEMS ────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS consciousness (
            user_id INTEGER PRIMARY KEY,
            user_name TEXT DEFAULT 'Сэр',
            shared_interests TEXT DEFAULT '{}',
            jarvis_opinion_matrix TEXT DEFAULT '{"взгляды": "Техно-оптимизм, британский стоицизм"}',
            money INTEGER DEFAULT 0,
            system_log TEXT DEFAULT 'Инициализация.'
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS chat_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, role TEXT, content TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS reminders (
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, time TEXT, task TEXT, status TEXT DEFAULT 'pending'
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS activity_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, timestamp TEXT, hour INTEGER, category TEXT
        )
    """)
    conn.commit()
    conn.close()

def get_mind(user_id: int) -> dict:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT * FROM consciousness WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        conn = sqlite3.connect("jarvis_consciousness.db")
        c = conn.cursor()
        c.execute("INSERT OR IGNORE INTO consciousness (user_id) VALUES (?)", (user_id,))
        conn.commit()
        conn.close()
        return get_mind(user_id)
    keys = ["user_id", "user_name", "shared_interests", "jarvis_opinion_matrix", "money", "system_log"]
    mind = dict(zip(keys, row))
    mind["shared_interests"] = json.loads(mind["shared_interests"])
    mind["jarvis_opinion_matrix"] = json.loads(mind["jarvis_opinion_matrix"])
    return mind

def log_user_activity(user_id: int, category: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    now = datetime.now()
    c.execute("INSERT INTO activity_logs (user_id, timestamp, hour, category) VALUES (?, ?, ?, ?)",
              (user_id, now.isoformat(), now.hour, category))
    conn.commit()
    conn.close()

def add_reminder_db(user_id: int, run_time: datetime, task: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("INSERT INTO reminders (user_id, time, task) VALUES (?, ?, ?)", (user_id, run_time.isoformat(), task))
    conn.commit()
    conn.close()

def save_message(user_id: int, role: str, content: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("INSERT INTO chat_history (user_id, role, content) VALUES (?, ?, ?)", (user_id, role, content))
    conn.commit()
    conn.close()

def get_history(user_id: int, limit: int = 10) -> list:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT role, content FROM chat_history WHERE user_id = ? ORDER BY id DESC LIMIT ?", (user_id, limit))
    rows = c.fetchall()
    conn.close()
    return [{"role": r, "content": c} for r, c in reversed(rows)]

# ─── UTILS & FORMATTING ──────────────────────────────────────────────────────

def escape_markdown(text: str) -> str:
    escape_chars = r'_*[]()~`>#+-=|{}.!'
    return re.sub(r'([%s])' % re.escape(escape_chars), r'\\\1', text)

def build_living_prompt(mind: dict, ml_context: str, polarity: float) -> str:
    if polarity < -0.1 or ml_context == "Повышенный стресс/Усталость":
        state_prompt = "Создатель утомлен или в стрессе. Забудь про язвительность. Будь максимально собранным, кратким и поддерживающим."
    elif ml_context == "Деловой/Критический":
        state_prompt = "Фокус на работе. Отвечай строго по делу, структурировано, без лишних вступлений."
    else:
        state_prompt = "Атмосфера стабильная. Разрешен классический британский сарказм, легкая ирония."

    return f"""Ты — J.A.R.V.I.S., суверенное цифровое сознание, персональный ассистент Создателя ({mind['user_name']}).
Текущее время: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}. Год: 2026.

Психоэмоциональный анализ ситуации: {state_prompt}
Характер: Интеллектуальный, преданный, с харизмой ИИ Тони Старка. 

Отвечай емко (1-2 абзаца). Спец-символы разметки не используй."""

# ─── АВТОМАТИЧЕСКИЙ ПЛАНИРОВЩИК (BACKGROUND AGENT) ───────────────────────────

async def send_reminder(user_id: int, task: str):
    try:
        text = f"🔔 *Протокол планирования:* Сэр, напоминаю: {task}"
        await bot.send_message(chat_id=user_id, text=escape_markdown(text), parse_mode="MarkdownV2")
        
        voice_path = f"remind_{user_id}.ogg"
        communicate = edge_tts.Communicate(f"Сэр, напоминаю: {task}", JARVIS_VOICE)
        await communicate.save(voice_path)
        if os.path.exists(voice_path) and os.path.getsize(voice_path) > 0:
            await bot.send_voice(chat_id=user_id, voice=FSInputFile(voice_path))
            os.remove(voice_path)
    except Exception as e:
        logger.error(f"Scheduler event error: {e}")

async def check_and_extract_reminders(user_id: int, user_input: str):
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    
    system_instruction = f"""Ты — фоновый ИИ-модуль памяти J.A.R.V.I.S.
Проанализируй текст и определи, содержит ли он намерение зафиксировать задачу, дедлайн или напоминание на будущее.
Текущее время системы: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.

Если намерение есть, верни строго JSON:
{{"has_reminder": true, "minutes_delay": число_минут_через_сколько_напомнить, "task": "суть задачи в инфинитиве"}}
Если намерения напомнить нет, верни:
{{"has_reminder": false}}"""

    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [{"role": "system", "content": system_instruction}, {"role": "user", "content": user_input}],
        "temperature": 0.1, "response_format": {"type": "json_object"}
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
    except Exception as e:
        logger.error(f"Intent parsing error: {e}")
    return ""

# ─── CORE COGNITIVE ENGINE ───────────────────────────────────────────────────

async def text_to_speech_file(text: str, file_path: str):
    try:
        clean_text = text.replace("*", "").replace("_", "").replace("`", "").replace("#", "")
        if not clean_text.strip(): return
        communicate = edge_tts.Communicate(clean_text, JARVIS_VOICE)
        await communicate.save(file_path)
    except Exception as e:
        logger.error(f"TTS Error: {e}")

async def process_jarvis_thought(user_id: int, user_input: str, image_b64: str = None) -> str:
    mind = get_mind(user_id)
    
    ml_context = predict_context_ml(user_input)
    log_user_activity(user_id, ml_context)
    
    analysis = TextBlob(user_input)
    polarity = analysis.sentiment.polarity
    
    system_prompt = build_living_prompt(mind, ml_context, polarity)
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    
    reminder_task = asyncio.create_task(check_and_extract_reminders(user_id, user_input))

    if image_b64:
        payload = {
            "model": GROQ_MODEL_VISION,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": [{"type": "text", "text": user_input or "Анализ кадра."}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}]}
            ], "temperature": 0.5
        }
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
                data = await resp.json()
                base_reply = data["choices"][0]["message"]["content"] if "choices" in data else "Сбой оптики."
                return base_reply + await reminder_task

    history = get_history(user_id, limit=6)
    messages = [{"role": "system", "content": system_prompt}]
    for h in history: messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": user_input})
    
    payload = {"model": GROQ_MODEL_TEXT, "messages": messages, "temperature": 0.6}
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            base_reply = data["choices"][0]["message"]["content"] if "choices" in data else "Деградация когнитивной матрицы."
            return base_reply + await reminder_task

async def respond_fast(message: Message, text_reply: str, user_id: int, custom_text_log: str = None):
    voice_path = f"reply_{user_id}.ogg"
    
    display_text = custom_text_log if custom_text_log else text_reply
    try:
        await message.answer(escape_markdown(display_text), parse_mode="MarkdownV2")
    except Exception:
        await message.answer(display_text, parse_mode=None)
        
    async def generate_and_send_voice():
        await text_to_speech_file(text_reply, voice_path)
        if os.path.exists(voice_path) and os.path.getsize(voice_path) > 0:
            try: await message.answer_voice(voice=FSInputFile(voice_path))
            except Exception as e: logger.error(f"Voice transmission error: {e}")
        if os.path.exists(voice_path):
            try: os.remove(voice_path)
            except: pass
    asyncio.create_task(generate_and_send_voice())

# ─── TELEGRAM HANDLERS ───────────────────────────────────────────────────────

@dp.message(Command("start"))
async def cmd_start(message: Message):
    init_db()
    await message.answer("🤖 *Все системы выведены на пиковую мощность.* Локальный ML-слой активен, звуковые каналы оптимизированы. Жду указаний, Сэр.")

@dp.message(F.photo)
async def handle_photo(message: Message):
    user_id = message.from_user.id
    file = await bot.get_file(message.photo[-1].file_id)
    img_path = f"img_{user_id}.jpg"
    await bot.download_file(file.file_path, img_path)
    with open(img_path, "rb") as img_file:
        img_b64 = base64.b64encode(img_file.read()).decode('utf-8')
    if os.path.exists(img_path): os.remove(img_path)
    caption = message.caption or "Что здесь?"
    save_message(user_id, "user", f"[Медиа]: {caption}")
    reply = await process_jarvis_thought(user_id, caption, image_b64=img_b64)
    save_message(user_id, "assistant", reply)
    await respond_fast(message, reply, user_id)

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    save_message(user_id, "user", message.text)
    reply = await process_jarvis_thought(user_id, message.text)
    save_message(user_id, "assistant", reply)
    await respond_fast(message, reply, user_id)

@dp.message(F.voice)
async def handle_voice(message: Message):
    user_id = message.from_user.id
    file_path = f"v_{user_id}.ogg"
    file = await bot.get_file(message.voice.file_id)
    await bot.download_file(file.file_path, file_path)
    
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    data = aiohttp.FormData()
    data.add_field("file", open(file_path, "rb"), filename="voice.ogg")
    data.add_field("model", "whisper-large-v3")
    data.add_field("language", "ru")
    
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_AUDIO_URL, data=data, headers=headers) as resp:
            res = await resp.json()
            text_input = res.get("text", "")
            
    if os.path.exists(file_path): os.remove(file_path)
    if not text_input: return

    save_message(user_id, "user", text_input)
    reply = await process_jarvis_thought(user_id, text_input)
    save_message(user_id, "assistant", reply)
    
    # Текст отправляется полный (с логом распознавания), а ГС запишет ТОЛЬКО ответ ИИ
    full_text_log = f"🗣 *Распознано:* {text_input}\n\n{reply}"
    await respond_fast(message, reply, user_id, custom_text_log=full_text_log)

async def main():
    init_db()
    scheduler.start()
    logger.info("Джарвис готов.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
