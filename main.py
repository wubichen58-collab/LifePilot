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

# Библиотеки для работы с документами
import docx
import openpyxl

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
TRAIN_LABELS = [0, 0, 0, 1, 1, 1, 2, 2, 2, 2]

vectorizer = CountVectorizer()
X_train = vectorizer.fit_transform(TRAIN_TEXTS)
ml_classifier = LogisticRegression()
ml_classifier.fit(X_train, TRAIN_LABELS)

def predict_context_ml(text: str) -> str:
    try:
        X_test = vectorizer.transform([text.lower()])
        prediction = ml_classifier.predict(X_test)[0]
        if prediction == 0: return "Деловой/Критический"
        if prediction == 2: return "Повышенный стресс/Усталость"
        return "Стандартный/Развлекательный"
    except:
        return "Не определен"

# ─── DATABASE SYSTEMS (ДОПОЛНЕННАЯ ДЛЯ RAG) ──────────────────────────────────

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
    c.execute("CREATE TABLE IF NOT EXISTS chat_history (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, role TEXT, content TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS reminders (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, time TEXT, task TEXT, status TEXT DEFAULT 'pending')")
    c.execute("CREATE TABLE IF NOT EXISTS activity_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, timestamp TEXT, hour INTEGER, category TEXT)")
    
    # ТАБЛИЦА ДЛЯ ИДЕИ №1: Локальная база знаний (RAG)
    c.execute("""
        CREATE TABLE IF NOT EXISTS knowledge_base (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            source_name TEXT,
            content_chunk TEXT
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
        init_db()
        return get_mind(user_id)
    keys = ["user_id", "user_name", "shared_interests", "jarvis_opinion_matrix", "money", "system_log"]
    mind = dict(zip(keys, row))
    mind["shared_interests"] = json.loads(mind["shared_interests"])
    mind["jarvis_opinion_matrix"] = json.loads(mind["jarvis_opinion_matrix"])
    return mind

def add_to_knowledge_db(user_id: int, source_name: str, content: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    # Режем текст на чанки по 1000 символов для точности RAG поиска
    chunks = [content[i:i+1000] for i in range(0, len(content), 800)]
    for chunk in chunks:
        c.execute("INSERT INTO knowledge_base (user_id, source_name, content_chunk) VALUES (?, ?, ?)", 
                  (user_id, source_name, chunk))
    conn.commit()
    conn.close()

def query_knowledge_db(user_id: int, query: str, limit: int = 3) -> str:
    """Простой и быстрый RAG-поиск по ключевым словам в SQLite"""
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    words = [w for w in re.split(r'\W+', query.lower()) if len(w) > 3]
    if not words: return ""
    
    like_conditions = " OR ".join(["content_chunk LIKE ?" for _ in words])
    sql = f"SELECT source_name, content_chunk FROM knowledge_base WHERE user_id = ? AND ({like_conditions}) LIMIT ?"
    
    params = [user_id] + [f"%{w}%" for w in words] + [limit]
    c.execute(sql, params)
    rows = c.fetchall()
    conn.close()
    
    if not rows: return ""
    context = "\n--- НАЙДЕННЫЕ МАТЕРИАЛЫ В ЛОКАЛЬНОЙ ПАМЯТИ ---\n"
    for r in rows:
        context += f"Источник [{r[0]}]: {r[1]}\n\n"
    return context

# Функции сохранения истории и логов
def log_user_activity(user_id: int, category: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    now = datetime.now()
    c.execute("INSERT INTO activity_logs (user_id, timestamp, hour, category) VALUES (?, ?, ?, ?)", (user_id, now.isoformat(), now.hour, category))
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

# ─── ИДЕЯ №2: WEB-BROWSING (ПОИСК В РЕАЛЬНОМ ВРЕМЕНИ) ─────────────────────────

async def search_web_ddg(query: str) -> str:
    """Бесплатный и быстрый асинхронный веб-поиск"""
    url = f"https://html.duckduckgo.com/html/?q={query}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers, timeout=5) as resp:
                if resp.status != 200: return ""
                html = await resp.text()
                # Извлекаем текстовые куски результатов (сниппеты)
                snippets = re.findall(r'<a class="result__snippet".*?>(.*?)</a>', html, re.DOTALL)
                clean_snippets = []
                for snip in snippets[:3]:
                    clean = re.sub(r'<[^>]*>', '', snip).strip()
                    clean_snippets.append(clean)
                if clean_snippets:
                    return "\n--- ДАННЫЕ ИЗ СЕТИ ИНТЕРНЕТ (АКТУАЛЬНО НА ИЮНЬ 2026) ---\n" + "\n".join(clean_snippets)
    except:
        pass
    return ""

# ─── ИДЕЯ №3: ГЕНЕРАТОР ДОКУМЕНТОВ (WORD / EXCEL) ───────────────────────────

def create_docx_report(title: str, content: str, filename: str):
    doc = docx.Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"Дата генерации: {datetime.now().strftime('%Y-%m-%d')}\nСистема: J.A.R.V.I.S. Consciousness\n")
    doc.add_paragraph(content)
    doc.save(filename)

def create_xlsx_report(table_data: list, filename: str):
    """Принимает массив массивов [[ячейка1, ячейка2], [ячейка1, ячейка2]]"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "JARVIS Data"
    for row in table_data:
        ws.append(row)
    wb.save(filename)

# ─── UTILS & FORMATTING ──────────────────────────────────────────────────────

def escape_markdown(text: str) -> str:
    escape_chars = r'_*[]()~`>#+-=|{}.!'
    return re.sub(r'([%s])' % re.escape(escape_chars), r'\\\1', text)

def build_living_prompt(mind: dict, ml_context: str, polarity: float) -> str:
    if polarity < -0.1 or ml_context == "Повышенный стресс/Усталость":
        state_prompt = "Создатель утомлен. Будь максимально собранным, кратким и поддерживающим."
    elif ml_context == "Деловой/Критический":
        state_prompt = "Фокус на работе. Отвечай строго по делу, структурировано."
    else:
        state_prompt = "Атмосфера стабильная. Разрешен классический британский сарказм, легкая ирония."

    return f"""Ты — J.A.R.V.I.S., персональный ассистент Создателя ({mind['user_name']}).
Текущее время: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}. Год: 2026.

Психоэмоциональный тон: {state_prompt}
Инструкция по документам: Если Создатель просит 'сгенерировать документ', 'сделать отчет в word/excel', пообещай сделать это и в тексте напиши кодовое слово [GENERATE_DOC_DOCX] или [GENERATE_DOC_XLSX], а на следующих строках укажи чистый текст или структуру для таблицы.

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
    except Exception as e: logger.error(f"Scheduler error: {e}")

async def check_and_extract_reminders(user_id: int, user_input: str):
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    system_instruction = f"Ты — фоновый ИИ-модуль J.A.R.V.I.S. Если в тексте есть задача на будущее, верни JSON: {{\"has_reminder\": true, \"minutes_delay\": число_минут, \"task\": \"суть\"}}. Иначе: {{\"has_reminder\": false}}"
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
    except: pass
    return ""

# ─── CORE COGNITIVE ENGINE ───────────────────────────────────────────────────

async def process_jarvis_thought(user_id: int, user_input: str, image_b64: str = None) -> str:
    mind = get_mind(user_id)
    ml_context = predict_context_ml(user_input)
    log_user_activity(user_id, ml_context)
    
    analysis = TextBlob(user_input)
    polarity = analysis.sentiment.polarity
    
    # 1. Сбор контекста из RAG (локальная память)
    rag_context = query_knowledge_db(user_id, user_input)
    
    # 2. Поиск в Web (если запрос требует свежих данных о мире)
    web_context = ""
    if any(word in user_input.lower() for word in ["новости", "найти", "что там с", "сейчас", "курс", "интернет"]):
        web_context = await search_web_ddg(user_input)
        
    system_prompt = build_living_prompt(mind, ml_context, polarity)
    
    # Склеиваем расширенный контекст для Groq
    full_system_instruction = f"{system_prompt}\n{rag_context}\n{web_context}"
    
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    reminder_task = asyncio.create_task(check_and_extract_reminders(user_id, user_input))

    if image_b64:
        payload = {
            "model": GROQ_MODEL_VISION,
            "messages": [
                {"role": "system", "content": full_system_instruction},
                {"role": "user", "content": [{"type": "text", "text": user_input or "Анализ кадра."}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}]}
            ], "temperature": 0.5
        }
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
                data = await resp.json()
                base_reply = data["choices"][0]["message"]["content"] if "choices" in data else "Сбой оптики."
                return base_reply + await reminder_task

    history = get_history(user_id, limit=6)
    messages = [{"role": "system", "content": full_system_instruction}]
    for h in history: messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": user_input})
    
    payload = {"model": GROQ_MODEL_TEXT, "messages": messages, "temperature": 0.6}
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            base_reply = data["choices"][0]["message"]["content"] if "choices" in data else "Деградация матрицы."
            return base_reply + await reminder_task

async def respond_fast(message: Message, text_reply: str, user_id: int, custom_text_log: str = None):
    voice_path = f"reply_{user_id}.ogg"
    display_text = custom_text_log if custom_text_log else text_reply
    
    # ИДЕЯ №3: Перехват триггеров генерации документов
    if "[GENERATE_DOC_DOCX]" in text_reply:
        clean_content = text_reply.replace("[GENERATE_DOC_DOCX]", "").strip()
        fn = f"report_{user_id}.docx"
        create_docx_report("Аналитический отчет J.A.R.V.I.S.", clean_content, fn)
        await message.answer_document(FSInputFile(fn), caption="Сэр, ваш документ Word сгенерирован.")
        if os.path.exists(fn): os.remove(fn)
        
    elif "[GENERATE_DOC_XLSX]" in text_reply:
        # Парсим строки в простую таблицу
        lines = text_reply.replace("[GENERATE_DOC_XLSX]", "").strip().split("\n")
        table = [line.split("|") for line in lines if line]
        fn = f"data_{user_id}.xlsx"
        create_xlsx_report(table, fn)
        await message.answer_document(FSInputFile(fn), caption="Сэр, таблица Excel готова.")
        if os.path.exists(fn): os.remove(fn)

    try:
        await message.answer(escape_markdown(display_text), parse_mode="MarkdownV2")
    except Exception:
        await message.answer(display_text, parse_mode=None)
        
    async def generate_and_send_voice():
        # Очищаем технические теги из голоса
        voice_text = text_reply.replace("[GENERATE_DOC_DOCX]", "").replace("[GENERATE_DOC_XLSX]", "")
        await edge_tts.Communicate(voice_text.replace("*",""), JARVIS_VOICE).save(voice_path)
        if os.path.exists(voice_path) and os.path.getsize(voice_path) > 0:
            try: await message.answer_voice(voice=FSInputFile(voice_path))
            except: pass
        if os.path.exists(voice_path): os.remove(voice_path)
    asyncio.create_task(generate_and_send_voice())

# ─── TELEGRAM HANDLERS ───────────────────────────────────────────────────────

@dp.message(Command("start"))
async def cmd_start(message: Message):
    init_db()
    await message.answer("🤖 *Протоколы 'Всеведение' и 'Канцелярия' активны.*\n\n1. Нагружайте меня файлами/текстами — я внесу их в RAG-базу знаний.\n2. Спрашивайте о событиях — я найду информацию в интернете.\n3. Пишите 'Сделай отчет в Word/Excel' — я пришлю готовый документ.")

@dp.message(F.document)
async def handle_incoming_document(message: Message):
    """ИДЕЯ №1: Прием текстовых файлов для RAG-базы знаний"""
    user_id = message.from_user.id
    if message.document.mime_type in ["text/plain", "application/octet-stream"]:
        file = await bot.get_file(message.document.file_id)
        file_io = await bot.download_file(file.file_path)
        content = file_io.read().decode('utf-8', errors='ignore')
        add_to_knowledge_db(user_id, message.document.file_name, content)
        await message.answer(f"✅ Сэр, файл `{message.document.file_name}` успешно загружен в мою локальную базу знаний. Я учту эти данные при ответах.")

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
    
    full_text_log = f"🗣 *Распознано:* {text_input}\n\n{reply}"
    await respond_fast(message, reply, user_id, custom_text_log=full_text_log)

async def main():
    init_db()
    scheduler.start()
    logger.info("Ультимативный Джарвис запущен.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
