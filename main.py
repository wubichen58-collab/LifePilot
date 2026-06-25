import asyncio
import logging
import os
import sqlite3
import json
import re
from datetime import datetime, timedelta
import aiohttp

import docx
import openpyxl
import pypdf
import edge_tts

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile
from apscheduler.schedulers.asyncio import AsyncIOScheduler

# ─── 1. ИНИЦИАЛИЗАЦИЯ И НАСТРОЙКИ ─────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
QWEATHER_KEY = os.environ.get("QWEATHER_KEY")

GROQ_CHAT_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL_TEXT = "llama-3.3-70b-versatile" 
GROQ_MODEL_FAST = "llama-3.1-8b-instant"
GROQ_MODEL_VISION = "llama-3.2-11b-vision-preview"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
scheduler = AsyncIOScheduler(timezone="Asia/Tokyo")

# ─── 2. ЛЕГКОЕ ХРАНИЛИЩЕ ДОКУМЕНТОВ (БЕЗ НАГРУЗКИ НА RAM) ─────────────────────
def add_to_document_store(user_id: int, source_name: str, content: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    # Фиксируем загрузку для авто-дайджеста
    conn.execute("INSERT INTO indexed_files (user_id, filename, timestamp) VALUES (?, ?, ?)", 
                 (user_id, source_name, datetime.now().isoformat()))
    # Записываем текст документа прямо в БД (потребление оперативной памяти = 0)
    conn.execute("INSERT INTO document_chunks (user_id, source_name, content) VALUES (?, ?, ?)",
                 (user_id, source_name, content))
    conn.commit()
    conn.close()

async def query_document_store(user_id: int, query: str) -> str:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    
    # Сверхбыстрый поиск совпадений по ключевым словам в SQLite
    words = [f"%{w}%" for w in query.split() if len(w) > 3]
    if not words:
        c.execute("SELECT source_name, content FROM document_chunks WHERE user_id = ? ORDER BY id DESC LIMIT 2", (user_id,))
    else:
        c.execute("SELECT source_name, content FROM document_chunks WHERE user_id = ? AND content LIKE ? LIMIT 3", (user_id, words[0]))
    
    rows = c.fetchall()
    conn.close()
    
    if not rows: return ""
    results = [f"[{row[0]}]: {row[1][:1200]}..." for row in rows]
    return "\n--- ДАННЫЕ ИЗ ВАШИХ ЛОКАЛЬНЫХ ФАЙЛОВ ---\n" + "\n".join(results)

# ─── 3. БАЗА ДАННЫХ И КОНТЕКСТ ПРОФИЛЯ ────────────────────────────────────────
def init_db():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS consciousness (user_id INTEGER PRIMARY KEY, user_name TEXT DEFAULT 'Сэр', shared_interests TEXT DEFAULT '{}', jarvis_opinion_matrix TEXT DEFAULT '{}', money INTEGER DEFAULT 0, system_log TEXT DEFAULT 'Инициализация.')""")
    c.execute("""CREATE TABLE IF NOT EXISTS chat_history (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, role TEXT, content TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS reminders (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, time TEXT, task TEXT, status TEXT DEFAULT 'pending')""")
    c.execute("""CREATE TABLE IF NOT EXISTS document_chunks (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, source_name TEXT, content TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS indexed_files (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, filename TEXT, timestamp TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS profile (key TEXT PRIMARY KEY, value TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS notes (id INTEGER PRIMARY KEY AUTOINCREMENT, tag TEXT, content TEXT)""")
    
    # Постоянный контекст для ваших текущих проектов
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

# ─── 4. ОБЛАЧНЫЕ ИИ-МЕТОДЫ (ОТВЕТ ЗА 1 СЕКУНДУ) ───────────────────────────────
async def call_groq_api(messages: list, model: str = GROQ_MODEL_TEXT, json_mode: bool = False) -> str:
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    payload = {"model": model, "messages": messages, "temperature": 0.3}
    if json_mode: payload["response_format"] = {"type": "json_object"}
    
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            return data["choices"][0]["message"]["content"]

async def analyze_emotion_via_api(text: str) -> str:
    """Заменяет локальный pipeline тяжелой нейросети на моментальный API запрос"""
    prompt = f"Определи эмоциональный контекст фразы: '{text}'. Ответь строго парой слов в формате 'Эмоция (Процент)', например: 'Усталость (90%)' или 'Нейтрально (100%)'."
    try:
        res = await call_groq_api([{"role": "user", "content": prompt}], model=GROQ_MODEL_FAST)
        return res.strip()
    except:
        return "Нейтрально"

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
    except: return "Данные о погоде временно недоступны."

# ─── 5. АГЕНТ ORCHESTRATOR И ЦЕНТРАЛЬНЫЙ ПРОЦЕССОР ────────────────────────────
async def process_jarvis_thought(user_id: int, user_input: str, image_b64: str = None) -> str:
    history = get_history(user_id)
    if user_input: save_message(user_id, "user", user_input)
    
    # Шаг 1: Агент Orchestrator (Декомпозиция сложной логики)
    orchestrator_prompt = f"""Ты — Архитектор систем J.A.R.V.I.S. Оцени запрос создателя: "{user_input}".
    Требуется ли для выполнения запроса сложный пошаговый план (написание научного эссе, финансовый анализ таблиц, составление графиков)?
    Верни строго JSON: {{"plan_needed": true, "steps": ["Шаг 1: ...", "Шаг 2: ..."]}} или {{"plan_needed": false}}.
    """
    
    plan_text = ""
    try:
        plan_resp = await call_groq_api([{"role": "system", "content": orchestrator_prompt}], model=GROQ_MODEL_FAST, json_mode=True)
        plan_data = json.loads(plan_resp)
        if plan_data.get("plan_needed"):
            steps_joined = "\n".join(plan_data.get("steps", []))
            plan_text = f"\n\n[ИНСТРУКЦИЯ ORCHESTRATOR ДЛЯ ВЫПОЛНЕНИЯ ПЛАНА]:\n{steps_joined}"
    except Exception as e:
        logger.error(f"Orchestrator error: {e}")

    # Шаг 2: Быстрый сбор контекста (Асинхронный запуск)
    emotion_task = asyncio.create_task(analyze_emotion_via_api(user_input))
    db_context_task = asyncio.create_task(query_document_store(user_id, user_input))
    emotion_label, doc_context = await asyncio.gather(emotion_task, db_context_task)
    
    # Шаг 3: Формирование системного промпта
    system_prompt = f"""Ты — ИИ-ассистент J.A.R.V.I.S.
    Текущая дата и время: {datetime.now().strftime('%Y-%m-%d %H:%M')}.
    Анализ состояния создателя: {emotion_label}.
    
    ПРОФИЛЬ СОЗДАТЕЛЯ И АКТУАЛЬНЫЕ ПРОЕКТЫ:
    {get_user_profile()}
    {doc_context}
    {plan_text}
    
    Отвечай лаконично, структурированно, в уважительном тоне ИИ-интеллекта."""

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

# ─── 6. ОБРАБОТКА ТЕЛЕГРАМ-СООБЩЕНИЙ И ДОКУМЕНТОВ ──────────────────────────────
@dp.message(F.document)
async def handle_document(message: Message):
    user_id = message.from_user.id
    init_db()
    file_name = message.document.file_name
    file_ext = file_name.split(".")[-1].lower() if "." in file_name else ""
    local_tmp_path = f"tmp_{user_id}_{file_name}"
    
    await bot.download_file((await bot.get_file(message.document.file_id)).file_path, local_tmp_path)
    content = ""
    await message.answer(f"⏳ Считываю структуру данных `{file_name}`...")
    
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
            add_to_document_store(user_id, file_name, content)
            await message.answer(f"✅ Документ `{file_name}` успешно сохранен в базу долговременной памяти.")
        else:
            await message.answer("⚠️ Текст в документе не обнаружен или файл пуст.")
    except Exception as e:
        logger.error(f"File indexing error: {e}")
        await message.answer(f"💥 Произошла ошибка при разборе структуры файла.")
    finally:
        if os.path.exists(local_tmp_path): os.remove(local_tmp_path)

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
        return await message.answer(f"💾 Информация зафиксирована под тегом: #{tag}")

    reply = await process_jarvis_thought(user_id, message.text)
    
    # Генерация озвучки через быстрое облако edge-tts
    voice_path = f"reply_{user_id}.ogg"
    try:
        dynamic_voice = get_voice_for_text(reply)
        await edge_tts.Communicate(reply.replace("*", "").replace("_", ""), dynamic_voice).save(voice_path)
        await message.answer(reply)
        if os.path.exists(voice_path):
            await message.answer_voice(voice=FSInputFile(voice_path))
            os.remove(voice_path)
    except Exception:
        await message.answer(reply)

# ─── 7. АВТОМАТИЧЕСКИЙ УТРЕННИЙ ДАЙДЖЕСТ (CRON) ────────────────────────────────
async def daily_briefing():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT DISTINCT user_id FROM consciousness")
    users = [row[0] for row in c.fetchall()]
    
    # Сбор данных по измененным или загруженным файлам за последние 24 часа
    yesterday = (datetime.now() - timedelta(days=1)).isoformat()
    c.execute("SELECT filename FROM indexed_files WHERE timestamp > ?", (yesterday,))
    recent_files = [r[0] for r in c.fetchall()]
    conn.close()

    weather = await get_weather_now()
    files_str = f"Загружено новых файлов за сутки: {len(recent_files)}." if recent_files else "Загрузки новых документов зафиксированы не были."
    
    text = f"Сэр, доброе утро. Сводный отчет по системам.\n{weather}\n{files_str}\nПотоки памяти стабильны. Готов к работе над вашими академическими проектами."
    
    for uid in users:
        try:
            await bot.send_message(uid, text)
            voice_path = f"brief_{uid}.ogg"
            await edge_tts.Communicate(text, "ru-RU-DmitryNeural").save(voice_path)
            await bot.send_voice(uid, FSInputFile(voice_path))
            os.remove(voice_path)
        except Exception as e: 
            logger.error(f"Briefing delivery failed: {e}")

# ─── 8. ЗАПУСК БОТА ───────────────────────────────────────────────────────────
async def main():
    init_db()
    # Ежедневная cron-задача на автоматический дайджест в 08:00 утра
    scheduler.add_job(daily_briefing, 'cron', hour=8, minute=0)
    scheduler.start()
    logger.info("J.A.R.V.I.S. переведен в оптимизированный режим мгновенного ответа.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
