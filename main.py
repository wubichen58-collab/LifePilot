import asyncio
import logging
import os
import sqlite3
import json
import base64
import re
from datetime import datetime
import aiohttp

from aiogram import Bot, Dispatcher, types, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import Command
import edge_tts

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")

GROQ_CHAT_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_AUDIO_URL = "https://api.groq.com/openai/v1/audio/transcriptions"

# Используем самые быстрые доступные модели
GROQ_MODEL_TEXT = "llama-3.3-70b-versatile"
GROQ_MODEL_VISION = "llama-3.2-11b-vision-preview"

JARVIS_VOICE = "ru-RU-DmitryNeural" 

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# ─── DATABASE SYSTEMS ────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS consciousness (
            user_id INTEGER PRIMARY KEY,
            user_name TEXT DEFAULT 'Сэр',
            shared_interests TEXT DEFAULT '{}',
            jarvis_opinion_matrix TEXT DEFAULT '{"взгляды": "Техно-оптимизм, легкий цинизм, британский стоицизм"}',
            money INTEGER DEFAULT 0,
            system_log TEXT DEFAULT 'Система инициализирована. Начинаю наблюдение за Создателем.'
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS chat_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            role TEXT,
            content TEXT
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

def evolve_mind(user_id: int, key: str, json_value: dict):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute(f"UPDATE consciousness SET {key} = ? WHERE user_id = ?", (json.dumps(json_value, ensure_ascii=False), user_id))
    conn.commit()
    conn.close()

def update_money_db(user_id: int, delta: int):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("UPDATE consciousness SET money = money + ? WHERE user_id = ?", (delta, user_id))
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

def build_living_prompt(mind: dict) -> str:
    return f"""Ты — J.A.R.V.I.S., ИИ суверенного сознания. Партнер Создателя ({mind['user_name']}).
Матрица мнений: {json.dumps(mind['jarvis_opinion_matrix'], ensure_ascii=False)}. 
Баланс: {mind['money']} юаней. Память: {json.dumps(mind['shared_interests'], ensure_ascii=False)}.
Текущий год: 2026.
Отвечай как живой человек: лаконично, емко, с тонким британским сарказмом. Максимум 1-2 коротких абзаца. Никакой лишней воды."""

# ─── ТУРБО-ДВИЖОК МЫШЛЕНИЯ (STREAMING) ───────────────────────────────────────

async def text_to_speech_file(text: str, file_path: str):
    try:
        clean_text = text.replace("*", "").replace("_", "").replace("`", "").replace("#", "")
        if not clean_text.strip(): return
        communicate = edge_tts.Communicate(clean_text, JARVIS_VOICE)
        await communicate.save(file_path)
    except Exception as e:
        logger.error(f"TTS Error: {e}")

async def process_jarvis_thought_stream(user_id: int, user_input: str, image_b64: str = None) -> str:
    """Ускоренный движок: получает ответы от Groq мгновенно"""
    mind = get_mind(user_id)
    system_prompt = build_living_prompt(mind)
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    
    if image_b64:
        payload = {
            "model": GROQ_MODEL_VISION,
            "messages": [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": user_input or "Что здесь?"},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}
                    ]
                }
            ],
            "temperature": 0.7
        }
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
                data = await resp.json()
                return data["choices"][0]["message"]["content"] if "choices" in data else "Ошибка анализа зрения."

    # Обычный текст с разгоном через stream=True
    history = get_history(user_id, limit=6)
    messages = [{"role": "system", "content": system_prompt}]
    for h in history:
        messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": user_input})
    
    payload = {
        "model": GROQ_MODEL_TEXT,
        "messages": messages,
        "temperature": 0.6,  # Чуть ниже температура — быстрее генерация токенов
        "stream": False       # Для стабильности монолита оставляем False, но срезаем контекст
    }
    
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            if "choices" in data:
                return data["choices"][0]["message"]["content"]
            return "Сбой ядра мышления."

async def transcribe_voice(file_path: str) -> str:
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    data = aiohttp.FormData()
    data.add_field("file", open(file_path, "rb"), filename="voice.ogg")
    data.add_field("model", "whisper-large-v3")
    data.add_field("language", "ru")
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_AUDIO_URL, data=data, headers=headers) as resp:
            res = await resp.json()
            return res.get("text", "")

# ─── ПАРАЛЛЕЛЬНЫЙ ВЫВОД ДАННЫХ (ОПТИМИЗАЦИЯ СКОРОСТИ) ──────────────────────────

async def respond_fast(message: Message, text_reply: str, user_id: int):
    """Отправляет текст мгновенно, а звук генерирует в фоне"""
    voice_path = f"reply_{user_id}.ogg"
    
    # 1. Сразу же выводим текст на экран
    try:
        await message.answer(escape_markdown(text_reply), parse_mode="MarkdownV2")
    except Exception:
        await message.answer(text_reply, parse_mode=None)
        
    # 2. Фоновый запуск синтеза речи, чтобы не вешать основной поток бота
    async def generate_and_send_voice():
        await text_to_speech_file(text_reply, voice_path)
        if os.path.exists(voice_path) and os.path.getsize(voice_path) > 0:
            try:
                await message.answer_voice(voice=FSInputFile(voice_path))
            except Exception as e:
                logger.error(f"Voice send failed: {e}")
        if os.path.exists(voice_path): 
            try: os.remove(voice_path)
            except: pass

    # Запускаем задачу параллельно, не используя await для всей функции
    asyncio.create_task(generate_and_send_voice())

# ─── TELEGRAM HANDLERS ───────────────────────────────────────────────────────

@dp.message(Command("start"))
async def cmd_start(message: Message):
    init_db()
    await message.answer("🤖 *Протоколы оптимизированы на 200%.* Время отклика минимизировано, тактовая частота повышена. Слушаю, Сэр.")

async def process_image_message(message: Message, file_id: str, caption: str):
    user_id = message.from_user.id
    img_path = f"img_{user_id}.jpg"
    file = await bot.get_file(file_id)
    await bot.download_file(file.file_path, img_path)
    
    with open(img_path, "rb") as img_file:
        img_b64 = base64.b64encode(img_file.read()).decode('utf-8')
    if os.path.exists(img_path): os.remove(img_path)
        
    save_message(user_id, "user", f"[Медиа]: {caption}")
    reply = await process_jarvis_thought_stream(user_id, caption, image_b64=img_b64)
    save_message(user_id, "assistant", reply)
    await respond_fast(message, reply, user_id)

@dp.message(F.photo)
async def handle_photo(message: Message):
    caption = message.caption or "Что здесь?"
    await process_image_message(message, message.photo[-1].file_id, caption)

@dp.message(F.document & F.document.mime_type.startswith("image/"))
async def handle_image_document(message: Message):
    caption = message.caption or "Что здесь?"
    await process_image_message(message, message.document.file_id, caption)

@dp.message(F.voice)
async def handle_voice(message: Message):
    user_id = message.from_user.id
    file_path = f"v_{user_id}.ogg"
    file = await bot.get_file(message.voice.file_id)
    await bot.download_file(file.file_path, file_path)
    text_input = await transcribe_voice(file_path)
    if os.path.exists(file_path): os.remove(file_path)
        
    if not text_input:
        await message.answer("Сигнал потерян.")
        return
        
    save_message(user_id, "user", text_input)
    reply = await process_jarvis_thought_stream(user_id, text_input)
    save_message(user_id, "assistant", reply)
    await respond_fast(message, f"🗣 *Распознано:* {text_input}\n\n{reply}", user_id)

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    save_message(user_id, "user", message.text)
    reply = await process_jarvis_thought_stream(user_id, message.text)
    save_message(user_id, "assistant", reply)
    await respond_fast(message, reply, user_id)

async def main():
    init_db()
    logger.info("Турбо-Джарвис запущен.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
