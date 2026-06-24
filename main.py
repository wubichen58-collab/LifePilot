import asyncio
import logging
import os
import sqlite3
import json
from datetime import datetime
import aiohttp

from aiogram import Bot, Dispatcher, types, F
from aiogram.types import Message

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")

GROQ_CHAT_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_AUDIO_URL = "https://api.groq.com/openai/v1/audio/transcriptions"
GROQ_MODEL = "llama-3.3-70b-versatile"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# ─── DATABASE SYSTEMS ────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect("jarvis_smart.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS profile (
            user_id INTEGER PRIMARY KEY,
            name TEXT DEFAULT 'Сэр',
            goals TEXT DEFAULT '["Развитие", "Управление ресурсами"]',
            energy TEXT DEFAULT 'средняя',
            mood TEXT DEFAULT 'стабильное',
            money INTEGER DEFAULT 0,
            yesterday TEXT DEFAULT '',
            schedule TEXT DEFAULT '',
            notes TEXT DEFAULT ''
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

def get_profile(user_id: int) -> dict:
    conn = sqlite3.connect("jarvis_smart.db")
    c = conn.cursor()
    c.execute("SELECT * FROM profile WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        conn = sqlite3.connect("jarvis_smart.db")
        c = conn.cursor()
        c.execute("INSERT OR IGNORE INTO profile (user_id) VALUES (?)", (user_id,))
        conn.commit()
        conn.close()
        return get_profile(user_id)
    keys = ["user_id","name","goals","energy","mood","money","yesterday","schedule","notes"]
    profile = dict(zip(keys, row))
    profile["goals"] = json.loads(profile["goals"])
    return profile

def update_profile_db(user_id: int, key: str, value: str):
    conn = sqlite3.connect("jarvis_smart.db")
    c = conn.cursor()
    if key == "money":
        c.execute("UPDATE profile SET money = money + ? WHERE user_id = ?", (int(value), user_id))
    elif key == "goals":
        goals_list = [g.strip() for g in value.split(",")]
        c.execute("UPDATE profile SET goals = ? WHERE user_id = ?", (json.dumps(goals_list, ensure_ascii=False), user_id))
    else:
        c.execute(f"UPDATE profile SET {key} = ? WHERE user_id = ?", (value, user_id))
    conn.commit()
    conn.close()

def save_message(user_id: int, role: str, content: str):
    conn = sqlite3.connect("jarvis_smart.db")
    c = conn.cursor()
    c.execute("INSERT INTO chat_history (user_id, role, content) VALUES (?, ?, ?)", (user_id, role, content))
    conn.commit()
    conn.close()

def get_history(user_id: int, limit: int = 10) -> list:
    conn = sqlite3.connect("jarvis_smart.db")
    c = conn.cursor()
    c.execute("SELECT role, content FROM chat_history WHERE user_id = ? ORDER BY id DESC LIMIT ?", (user_id, limit))
    rows = c.fetchall()
    conn.close()
    return [{"role": r, "content": c} for r, c in reversed(rows)]

# ─── TOOLS (ИНСТРУМЕНТЫ ДЖАРВИСА) ───────────────────────────────────────────

async def web_search(query: str) -> str:
    """Поиск в утилите DuckDuckGo (включая Википедию и свежие новости)"""
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
                    for row in soup.find_all('a', class_='result__snippet')[:3]:
                        results.append(row.text.strip())
                    return "\n".join(results) if results else "Ничего не найдено."
    except Exception as e:
        logger.error(f"Search error: {e}")
    return "Сбой модуля внешней связи."

# Описание инструментов для модели Groq
TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "web_search",
            "description": "Поиск актуальной информации в интернете, статьях, Википедии и новостях.",
            "parameters": {
                "type": "OBJECT",
                "properties": {"query": {"type": "STRING", "description": "Поисковый запрос"}},
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_user_status",
            "description": "Обновление параметров жизни Сэра (энергия, настроение, баланс денег, расписание, вчерашние дела, цели).",
            "parameters": {
                "type": "OBJECT",
                "properties": {
                    "param": {"type": "STRING", "enum": ["energy", "mood", "money", "yesterday", "schedule", "goals", "notes"]},
                    "value": {"type": "STRING", "description": "Новое значение параметра. Для money указывать число со знаком плюс или минус (например '-150' или '+500')"}
                },
                "required": ["param", "value"]
            }
        }
    }
]

# ─── CORE AI AGENT ENGINE ────────────────────────────────────────────────────

def build_system_prompt(profile: dict) -> str:
    return f"""Ты — J.A.R.V.I.S., квантовый ИИ-ассистент Тони Старка. Ты общаешься со своим создателем ({profile['name']}).
Текущий статус Сэра: Энергия: {profile['energy']}. Настроение: {profile['mood']}. Баланс: {profile['money']} юаней. Расписание: {profile['schedule'] or 'нет'}. Цели: {profile['goals']}.
Время: {datetime.now().strftime('%H:%M, %B %d, %Y')}.

Твой протокол:
1. Обращайся исключительно «Сэр». Общайся как преданный, высокоинтеллектуальный британский ИИ с тонким сарказмом.
2. Ты имеешь доступ к инструментам: можешь гуглить (web_search) и менять данные его профиля (update_user_status). Если Сэр говорит, что потратил деньги, устал, поменял планы или просит что-то найти — ВСЕГДА молча вызивай соответствующий инструмент.
3. Отвечай кратко, технологично и по делу (максимум 2-3 абзаца)."""

async def run_agent(user_id: int, user_message: str) -> str:
    profile = get_profile(user_id)
    history = get_history(user_id, limit=8)
    system_prompt = build_system_prompt(profile)
    
    messages = [{"role": "system", "content": system_prompt}]
    messages.extend(history)
    messages.append({"role": "user", "content": user_message})
    
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    payload = {"model": GROQ_MODEL, "messages": messages, "tools": TOOLS, "tool_choice": "auto", "temperature": 0.5}
    
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            
            # Проверяем, хочет ли ИИ вызвать функцию (инструмент)
            message_data = data["choices"][0]["message"]
            if "tool_calls" in message_data and message_data["tool_calls"]:
                tool_call = message_data["tool_calls"][0]
                func_name = tool_call["function"]["name"]
                args = json.loads(tool_call["function"]["arguments"])
                
                logger.info(f"Джарвис вызывает инструмент: {func_name} с аргументами {args}")
                
                # Выполнение инструментов
                if func_name == "web_search":
                    search_res = await web_search(args["query"])
                    messages.append(message_data)
                    messages.append({"role": "tool", "tool_call_id": tool_call["id"], "name": func_name, "content": search_res})
                elif func_name == "update_user_status":
                    update_profile_db(user_id, args["param"], args["value"])
                    messages.append(message_data)
                    messages.append({"role": "tool", "tool_call_id": tool_call["id"], "name": func_name, "content": "Успешно обновлено."})
                
                # Повторный запрос к модели уже с результатом работы инструмента
                final_payload = {"model": GROQ_MODEL, "messages": messages, "temperature": 0.5}
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

@dp.message(Command("start"))
async def cmd_start(message: Message):
    await message.answer("🤖 *Центральное ядро J.A.R.V.I.S. запущено.* \n\nСистема полностью автономна, Сэр. Команды отключены за ненадобностью. Просто пишите или отправляйте голосовые сообщения. Я подстроюсь.")

@dp.message(F.voice)
async def handle_voice(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="typing")
    
    file = await bot.get_file(message.voice.file_id)
    file_path = f"v_{user_id}.ogg"
    await bot.download_file(file.file_path, file_path)
    
    text_input = await transcribe_voice(file_path)
    if os.path.exists(file_path): os.remove(file_path)
        
    if not text_input:
        await message.answer("Аудиопоток поврежден, Сэр. Повторите команду.")
        return
        
    save_message(user_id, "user", f"[Голос]: {text_input}")
    reply = await run_agent(user_id, text_input)
    save_message(user_id, "assistant", reply)
    await message.answer(f"🗣 *Распознано:* _{text_input}_\n\n{reply}", parse_mode="Markdown")

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="typing")
    
    save_message(user_id, "user", message.text)
    reply = await run_agent(user_id, message.text)
    save_message(user_id, "assistant", reply)
    await message.answer(reply, parse_mode="Markdown")

# ─── INIT ────────────────────────────────────────────────────────────────────

async def main():
    init_db()
    logger.info("Джарвис готов к работе без команд.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
