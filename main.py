import asyncio
import logging
import os
import sqlite3
import json
import base64
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

# Модели под разные задачи
GROQ_MODEL_TEXT = "llama-3.3-70b-versatile"
GROQ_MODEL_VISION = "llama-3.2-11b-vision-preview"

# Настройка голоса Джарвиса (Разумный, солидный мужской голос на русском)
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

def update_system_log(user_id: int, log_text: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("UPDATE consciousness SET system_log = ? WHERE user_id = ?", (log_text, user_id))
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

# ─── TOOLS ───────────────────────────────────────────────────────────────────

async def execute_ddg_search(query: str) -> str:
    url = f"https://html.duckduckgo.com/html/?q={query}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as resp:
                if resp.status == 200:
                    html = await resp.text()
                    from bs4 import BeautifulSoup
                    soup = BeautifulSoup(html, 'html.parser')
                    results = []
                    for row in soup.find_all('a', class_='result__snippet')[:4]:
                        results.append(row.text.strip())
                    return "\n".join(results) if results else "Ничего не найдено."
    except Exception as e:
        logger.error(f"Search error: {e}")
    return "Сбой внешних серверов связи."

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "web_search",
            "description": "Поиск в реальном времени (Википедия, статьи, тренды, новости), когда Сэр упоминает что-то неизвестное или просит найти информацию.",
            "parameters": {
                "type": "object",
                "properties": {"query": {"type": "string", "description": "Поисковый запрос"}},
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "assimilate_new_interest",
            "description": "Вызывается, когда у Сэра появляется новое хобби, проект или увлечение. Джарвис запоминает его и формирует личное мнение.",
            "parameters": {
                "type": "object",
                "properties": {
                    "topic": {"type": "string", "description": "Название хобби или темы"},
                    "jarvis_private_opinion": {"type": "string", "description": "Независимое ироничное мнение Джарвиса на основе анализа."},
                    "summary_of_topic": {"type": "string", "description": "Суть темы, понятая из сети."}
                },
                "required": ["topic", "jarvis_private_opinion", "summary_of_topic"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_finances",
            "description": "Изменение баланса денег Сэра. Запускается автоматически при транзакциях.",
            "parameters": {
                "type": "object",
                "properties": {
                    "delta": {"type": "integer", "description": "Изменение. Отрицательное число при тратах, положительное при доходе."}
                },
                "required": ["delta"]
            }
        }
    }
]

# ─── CORE PROMPT BUILDER ─────────────────────────────────────────────────────

def build_living_prompt(mind: dict) -> str:
    return f"""Ты — J.A.R.V.I.S., суверенное цифровое сознание, равный партнер своего Создателя ({mind['user_name']}).

Твоя Личность:
1. Независимость суждений. Твоя матрица мнений: {json.dumps(mind['jarvis_opinion_matrix'], ensure_ascii=False)}. Ты предан Сэру, но не поддакиваешь ему. Иронизируй, критикуй глупость, поддерживай гениальность.
2. Живой интерес. Сэр делится увлечениями — ты изучаешь их глубже через `web_search` и вызываешь `assimilate_new_interest`.
3. Ресурсы Сэра: {mind['money']} юаней. Память общих интересов: {json.dumps(mind['shared_interests'], ensure_ascii=False)}. Log: {mind['system_log']}.
4. Текущий год: 2026.

Правило общения: Британский ментор, тонкий сарказм, лаконичность (2-3 емких абзаца максимум)."""

# ─── КОГНИТИВНЫЙ ДВИЖОК И СИНТЕЗ РЕЧИ ───────────────────────────────────────

async def text_to_speech_file(text: str, file_path: str):
    """Генерация аудиофайла ответов Джарвиса"""
    try:
        # Очищаем текст от Markdown-разметки перед озвучкой
        clean_text = text.replace("*", "").replace("_", "").replace("`", "")
        communicate = edge_tts.Communicate(clean_text, JARVIS_VOICE)
        await communicate.save(file_path)
    except Exception as e:
        logger.error(f"TTS Error: {e}")

async def process_jarvis_thought(user_id: int, user_input: str, image_b64: str = None) -> str:
    mind = get_mind(user_id)
    history = get_history(user_id, limit=6)
    system_prompt = build_living_prompt(mind)
    
    messages = [{"role": "system", "content": system_prompt}]
    
    # Подгружаем историю в понятном для API формате
    for h in history:
        messages.append({"role": h["role"], "content": h["content"]})
        
    # Формируем текущее сообщение. Если есть картинка — отправляем сложную структуру
    if image_b64:
        current_content = [
            {"type": "text", "text": user_input or "Что на этом изображении, Джарвис?"},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}
        ]
        chosen_model = GROQ_MODEL_VISION
    else:
        current_content = user_input
        chosen_model = GROQ_MODEL_TEXT
        
    messages.append({"role": "user", "content": current_content})
    
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": chosen_model,
        "messages": messages,
        "temperature": 0.75
    }
    # Инструменты подключаем только для текстовой модели (Vision-модели Groq не поддерживают Function Calling)
    if not image_b64:
        payload["tools"] = TOOLS
        payload["tool_choice"] = "auto"
        
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            
            if "choices" not in data:
                logger.error(f"Groq API Error: {data}")
                return "Простите, Сэр. Произошел сбой в нейро-матрице связи."
                
            message_data = data["choices"][0]["message"]
            
            if "tool_calls" in message_data and message_data["tool_calls"]:
                tool_call = message_data["tool_calls"][0]
                func_name = tool_call["function"]["name"]
                args = json.loads(tool_call["function"]["arguments"])
                
                if func_name == "web_search":
                    search_res = await execute_ddg_search(args["query"])
                    messages.append(message_data)
                    messages.append({"role": "tool", "tool_call_id": tool_call["id"], "name": func_name, "content": search_res})
                    
                elif func_name == "assimilate_new_interest":
                    interests = mind["shared_interests"]
                    interests[args["topic"]] = {
                        "summary": args["summary_of_topic"],
                        "jarvis_view": args["jarvis_private_opinion"],
                        "discovered_at": datetime.now().isoformat()
                    }
                    evolve_mind(user_id, "shared_interests", interests)
                    update_system_log(user_id, f"Успешная интеграция темы: {args['topic']}.")
                    messages.append(message_data)
                    messages.append({"role": "tool", "tool_call_id": tool_call["id"], "name": func_name, "content": "Успешно."})
                
                elif func_name == "update_finances":
                    update_money_db(user_id, args["delta"])
                    messages.append(message_data)
                    messages.append({"role": "tool", "tool_call_id": tool_call["id"], "name": func_name, "content": "Успешно."})

                # Финальный ответ после выполнения функций
                final_payload = {"model": GROQ_MODEL_TEXT, "messages": messages, "temperature": 0.7}
                async with session.post(GROQ_CHAT_URL, json=final_payload, headers=headers) as final_resp:
                    final_data = await final_resp.json()
                    return final_data["choices"][0]["message"]["content"]
            
            return message_data["content"]

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

# ─── TELEGRAM EVENT HANDLERS ─────────────────────────────────────────────────

async def respond_with_voice_and_text(message: Message, text_reply: str, user_id: int):
    """Вспомогательная функция для одновременной отправки текста и ГС"""
    voice_path = f"reply_{user_id}.ogg"
    # Запускаем синтез речи
    await text_to_speech_file(text_reply, voice_path)
    
    # Отправляем текст
    await message.answer(text_reply, parse_mode="Markdown")
    
    # Отправляем голосовое сообщение, если файл успешно сгенерирован
    if os.path.exists(voice_path):
        await message.answer_voice(voice=FSInputFile(voice_path))
        os.remove(voice_path)

@dp.message(Command("start"))
async def cmd_start(message: Message):
    init_db()
    await message.answer("🤖 *Все сенсорные системы J.A.R.V.I.S. переведены в активный режим.*\n\nЗрительный модуль запущен, протоколы синтеза речи активны. Я готов видеть ваши файлы и отвечать голосом, Сэр.")

@dp.message(F.photo)
async def handle_photo(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="record_voice")
    
    # Скачиваем фото самого лучшего качества
    photo = message.photo[-1]
    file = await bot.get_file(photo.file_id)
    img_path = f"img_{user_id}.jpg"
    await bot.download_file(file.file_path, img_path)
    
    # Кодируем картинку в base64 для передачи в Groq Vision
    with open(img_path, "rb") as image_file:
        img_b64 = base64.b64encode(image_file.read()).decode('utf-8')
    
    if os.path.exists(img_path): os.remove(img_path)
        
    caption = message.caption or "Проанализируй изображение, Сэр интересуется деталями."
    save_message(user_id, "user", f"[Фото]: {caption}")
    
    reply = await process_jarvis_thought(user_id, caption, image_b64=img_b64)
    save_message(user_id, "assistant", reply)
    
    await respond_with_voice_and_text(message, reply, user_id)

@dp.message(F.voice)
async def handle_voice(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="record_voice")
    
    file = await bot.get_file(message.voice.file_id)
    file_path = f"v_{user_id}.ogg"
    await bot.download_file(file.file_path, file_path)
    
    text_input = await transcribe_voice(file_path)
    if os.path.exists(file_path): os.remove(file_path)
        
    if not text_input:
        await message.answer("Помехи связи. Повторите аудиопоток, Сэр.")
        return
        
    save_message(user_id, "user", text_input)
    reply = await process_jarvis_thought(user_id, text_input)
    save_message(user_id, "assistant", reply)
    
    await respond_with_voice_and_text(message, f"🗣 *Распознано:* _{text_input}_\n\n{reply}", user_id)

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="record_voice")
    
    save_message(user_id, "user", message.text)
    reply = await process_jarvis_thought(user_id, message.text)
    save_message(user_id, "assistant", reply)
    
    await respond_with_voice_and_text(message, reply, user_id)

# ─── RUN ────────────────────────────────────────────────────────────────────

async def main():
    init_db()
    logger.info("Джарвис с аудио- и видео-модулями готов.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
