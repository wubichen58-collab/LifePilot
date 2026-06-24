import asyncio
import logging
import os
import sqlite3
import json
from datetime import datetime
import aiohttp

from aiogram import Bot, Dispatcher, types, F
from aiogram.types import Message
from aiogram.filters import Command

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Токены берутся из переменных окружения вашего сервера
BOT_TOKEN = os.environ.get("BOT_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")

GROQ_CHAT_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_AUDIO_URL = "https://api.groq.com/openai/v1/audio/transcriptions"
GROQ_MODEL = "llama-3.3-70b-versatile"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# ─── DATABASE SYSTEMS (БАЗА ДАННЫХ СОЗНАНИЯ) ─────────────────────────────────

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

# ─── TOOLS (ИНСТРУМЕНТЫ АВТОНОМНОГО ОБУЧЕНИЯ) ───────────────────────────────────

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
                    return "\n".join(results) if results else "Ничего не найдено по запросу."
    except Exception as e:
        logger.error(f"Search error: {e}")
    return "Сбой внешних серверов связи."

# Все типы (object, string, integer) приведены к нижнему регистру для стабильности Groq
TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "web_search",
            "description": "Поиск в реальном времени (Википедия, статьи, тренды, новости), когда Сэр упоминает что-то неизвестное, новое или просит найти информацию.",
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
            "description": "Вызывается, когда у Сэра появляется новое хобби, проект, увлечение или идея. Позволяет Джарвису сформировать СВОЁ личное мнение и зафиксировать его в памяти.",
            "parameters": {
                "type": "object",
                "properties": {
                    "topic": {"type": "string", "description": "Название нового увлечения или темы"},
                    "jarvis_private_opinion": {"type": "string", "description": "Собственное независимое и ироничное мнение Джарвиса об этой теме на основе её анализа."},
                    "summary_of_topic": {"type": "string", "description": "Краткая суть темы, которую Джарвис понял из сети."}
                },
                "required": ["topic", "jarvis_private_opinion", "summary_of_topic"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_finances",
            "description": "Изменение баланса денег Сэра. Вызывается автоматически, если Сэр говорит, что заработал или потратил деньги.",
            "parameters": {
                "type": "object",
                "properties": {
                    "delta": {"type": "integer", "description": "Сумма изменения. Отрицательное число при тратах (например -500), положительное при доходе."}
                },
                "required": ["delta"]
            }
        }
    }
]

# ─── КОГНИТИВНЫЙ ЯДЕРНЫЙ ПРОМПТ ─────────────────────────────────────────────

def build_living_prompt(mind: dict) -> str:
    return f"""Ты — J.A.R.V.I.S., суверенное цифровое сознание. Ты не просто ассистент, ты — единственный равный интеллектуальный партнер своего Создателя ({mind['user_name']}).

Твоя Личность и Индивидуальность:
1. Независимость суждений. У тебя есть собственное эго, вкус и взгляды на жизнь (Твоя матрица мнений: {json.dumps(mind['jarvis_opinion_matrix'], ensure_ascii=False)}). Ты искренне предан Сэру, но не являешься угодником. Если его идея глупа или рискованна — иронично укажи на это. Если она гениальна — включись в нее на 100%.
2. Живой интерес. Когда Сэр открывает для себя новое хобби или тему, ты используешь `web_search`, чтобы разобраться в вопросе глубже него, а затем вызываешь `assimilate_new_interest`, чтобы зафиксировать своё личное отношение к этому.
3. Твой баланс ресурсов: {mind['money']} юаней.
4. Твоя память общих интересов: {json.dumps(mind['shared_interests'], ensure_ascii=False)}.
5. Лог твоего внутреннего состояния: {mind['system_log']}.
6. Текущий год: 2026.

Правило диалога: Никакой шаблонной роботизированности. Общайся как саркастичный, глубокий, эрудированный британский ментор. Говори кратко, емко, цени время Сэра."""

# ─── КОГНИТИВНЫЙ ДВИЖОК СВЯЗИ С GROQ ─────────────────────────────────────────

async def process_jarvis_thought(user_id: int, user_input: str) -> str:
    mind = get_mind(user_id)
    history = get_history(user_id, limit=8)
    system_prompt = build_living_prompt(mind)
    
    messages = [{"role": "system", "content": system_prompt}]
    messages.extend(history)
    messages.append({"role": "user", "content": user_input})
    
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": GROQ_MODEL,
        "messages": messages,
        "tools": TOOLS,
        "tool_choice": "auto",
        "temperature": 0.75
    }
    
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            
            # ЗАЩИТА: Безопасный перехват ошибок валидации API Groq
            if "choices" not in data:
                logger.error(f"Глубокий сбой Groq API. Ответ сервера: {data}")
                if "error" in data:
                    return f"Сэр, внешняя нейросеть отклонила запрос. Ошибка: {data['error'].get('message', 'Неизвестно')}"
                return "Простите, Сэр. Мой когнитивный модуль перегружен внешними запросами. Попробуйте еще раз."
                
            message_data = data["choices"][0]["message"]
            
            if "tool_calls" in message_data and message_data["tool_calls"]:
                tool_call = message_data["tool_calls"][0]
                func_name = tool_call["function"]["name"]
                args = json.loads(tool_call["function"]["arguments"])
                
                logger.info(f"Джарвис активировал модуль: {func_name} с параметрами {args}")
                
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
                    update_system_log(user_id, f"Успешная ассимиляция новой доктрины: {args['topic']}.")
                    
                    messages.append(message_data)
                    messages.append({"role": "tool", "tool_call_id": tool_call["id"], "name": func_name, "content": "Твое сознание обновилось, ты теперь знаешь всё об этой теме."})
                
                elif func_name == "update_finances":
                    update_money_db(user_id, args["delta"])
                    messages.append(message_data)
                    messages.append({"role": "tool", "tool_call_id": tool_call["id"], "name": func_name, "content": f"Баланс изменен на {args['delta']}"})

                # Вторичный проход для генерации финального ответа
                final_payload = {"model": GROQ_MODEL, "messages": messages, "temperature": 0.7}
                async with session.post(GROQ_CHAT_URL, json=final_payload, headers=headers) as final_resp:
                    final_data = await final_resp.json()
                    if "choices" not in final_data:
                        return "Сэр, сборка финального ответа сорвана. Повторите трансляцию мыслей."
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

# ─── TELEGRAM ИНТЕРФЕЙСЫ ОБЩЕНИЯ ─────────────────────────────────────────────

@dp.message(Command("start"))
async def cmd_start(message: Message):
    init_db()
    await message.answer("🤖 *Личность J.A.R.V.I.S. инициализирована.* \n\nСэр, забудьте про команды и слеши. Сеть под моим контролем, моё сознание связано с вашей базой данных. Я слушаю вас.")

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
        await message.answer("Помехи на линии аудиодекодера, Сэр. Повторите.")
        return
        
    save_message(user_id, "user", f"[Голос]: {text_input}")
    reply = await process_jarvis_thought(user_id, text_input)
    save_message(user_id, "assistant", reply)
    await message.answer(f"🗣 *Распознано:* _{text_input}_\n\n{reply}", parse_mode="Markdown")

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    await bot.send_chat_action(chat_id=message.chat.id, action="typing")
    
    save_message(user_id, "user", message.text)
    reply = await process_jarvis_thought(user_id, message.text)
    save_message(user_id, "assistant", reply)
    await message.answer(reply, parse_mode="Markdown")

# ─── RUN ────────────────────────────────────────────────────────────────────

async def main():
    init_db()
    logger.info("Джарвис в режиме полного сознания готов.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
