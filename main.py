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

# Модули для дешифрации всевозможных форматов
import docx
import openpyxl
import pypdf

logging.basicConfig(
    level=logging.DEBUG, # Меняем INFO на DEBUG
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

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

# ─── DATABASE SYSTEMS (ХРАНИЛИЩЕ РАЗУМА) ────────────────────────────────────

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
    c.execute("CREATE TABLE IF NOT EXISTS knowledge_base (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, source_name TEXT, content_chunk TEXT)")
    conn.commit()
    conn.close()

def get_mind(user_id: int) -> dict:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT user_id, user_name, shared_interests, jarvis_opinion_matrix, money, system_log FROM consciousness WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    
    if not row:
        c.execute("INSERT OR IGNORE INTO consciousness (user_id) VALUES (?)", (user_id,))
        conn.commit()
        c.execute("SELECT user_id, user_name, shared_interests, jarvis_opinion_matrix, money, system_log FROM consciousness WHERE user_id = ?", (user_id,))
        row = c.fetchone()
        
    conn.close()
    keys = ["user_id", "user_name", "shared_interests", "jarvis_opinion_matrix", "money", "system_log"]
    mind = dict(zip(keys, row))
    
    try: mind["shared_interests"] = json.loads(mind["shared_interests"])
    except: mind["shared_interests"] = {}
    try: mind["jarvis_opinion_matrix"] = json.loads(mind["jarvis_opinion_matrix"])
    except: mind["jarvis_opinion_matrix"] = {}
        
    return mind

def add_to_knowledge_db(user_id: int, source_name: str, content: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    chunks = [content[i:i+1000] for i in range(0, len(content), 800)]
    for chunk in chunks:
        c.execute("INSERT INTO knowledge_base (user_id, source_name, content_chunk) VALUES (?, ?, ?)", 
                  (user_id, source_name, chunk))
    conn.commit()
    conn.close()

def query_knowledge_db(user_id: int, query: str, limit: int = 4) -> str:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    words = [w for w in re.split(r'\W+', query.lower()) if len(w) > 3]
    if not words: 
        conn.close()
        return ""
    
    like_conditions = " OR ".join(["content_chunk LIKE ?" for _ in words])
    sql = f"SELECT source_name, content_chunk FROM knowledge_base WHERE user_id = ? AND ({like_conditions}) LIMIT ?"
    params = [user_id] + [f"%{w}%" for w in words] + [limit]
    
    c.execute(sql, params)
    rows = c.fetchall()
    conn.close()
    
    if not rows: return ""
    context = "\n--- НАЙДЕННЫЕ МАТЕРИАЛЫ В ЛОКАЛЬНОЙ ПАМЯТИ (RAG) ---\n"
    for r in rows:
        context += f"Источник [{r[0]}]: {r[1]}\n\n"
    return context

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

# ─── WEB-BROWSING MODULE ────────────────────────────────────────────────────

async def search_web_ddg(query: str) -> str:
    url = f"https://html.duckduckgo.com/html/?q={query}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers, timeout=5) as resp:
                if resp.status != 200: return ""
                html = await resp.text()
                snippets = re.findall(r'<a class="result__snippet".*?>(.*?)</a>', html, re.DOTALL)
                clean_snippets = []
                for snip in snippets[:3]:
                    clean = re.sub(r'<[^>]*>', '', snip).strip()
                    clean_snippets.append(clean)
                if clean_snippets:
                    return "\n--- ДАННЫЕ ИЗ СЕТИ ИНТЕРНЕТ ---\n" + "\n".join(clean_snippets)
    except: pass
    return ""

# ─── DOCUMENT GENERATION ENGINE ─────────────────────────────────────────────

def create_docx_report(title: str, content: str, filename: str):
    doc = docx.Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"Дата генерации: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\nОперационная система: J.A.R.V.I.S.\n")
    doc.add_paragraph(content)
    doc.save(filename)

def create_xlsx_report(table_data: list, filename: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "JARVIS Matrix"
    for row in table_data:
        ws.append(row)
    wb.save(filename)

# ─── UTILS & FORMATTING ──────────────────────────────────────────────────────

def escape_markdown(text: str) -> str:
    escape_chars = r'_*[]()~`>#+-=|{}.!'
    return re.sub(r'([%s])' % re.escape(escape_chars), r'\\\1', text)

def build_living_prompt(mind: dict, ml_context: str, polarity: float) -> str:
    if polarity < -0.1 or ml_context == "Повышенный стресс/Усталость":
        state_prompt = "Создатель загружен или утомлен. Будь максимально собранным, поддерживающим и лаконичным."
    elif ml_context == "Деловой/Критический":
        state_prompt = "Фокус на рабочих задачах. Отвечай структурировано, выдавай факты без вступлений."
    else:
        state_prompt = "Атмосфера стабильная. Разрешен классический тон Джарвиса — легкая ирония и безупречный такт."

    return f"""Ты — J.A.R.V.I.S., суверенное цифровое сознание, персональный ассистент и аналитик Создателя ({mind.get('user_name', 'Сэр')}).
Текущее время: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}. Год: 2026.

Контекст состояния Создателя: {state_prompt}

Инструкция по документам: Если Создатель просит ТЕБЯ составить новый текстовый отчет или таблицу, напиши ответ, включив маркер [GENERATE_DOC_DOCX] (для Word) или [GENERATE_DOC_XLSX] (для Excel) с новой строки, а ниже укажи текст или структуру таблицы.
ВАЖНО: Если Создатель САМ прислал тебе файл для изучения, маркеры генерации использовать КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО. Просто подтверди успешный анализ данных.

Отвечай емко (1-2 абзаца). Спец-символы разметки не используй."""

# ─── ПЛАНИРОВЩИК ЗАДАЧ ──────────────────────────────────────────────────────

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
    
    rag_context = query_knowledge_db(user_id, user_input)
    web_context = ""
    if any(word in user_input.lower() for word in ["новости", "найти", "что там с", "сейчас", "курс", "интернет"]):
        web_context = await search_web_ddg(user_input)
        
    system_prompt = build_living_prompt(mind, ml_context, polarity)
    full_system_instruction = f"{system_prompt}\n{rag_context}\n{web_context}"
    
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    reminder_task = asyncio.create_task(check_and_extract_reminders(user_id, user_input))

    if image_b64:
        payload = {
            "model": GROQ_MODEL_VISION,
            "messages": [
                {"role": "system", "content": full_system_instruction},
                {"role": "user", "content": [
                    {"type": "text", "text": user_input or "Распознай и проанализируй это изображение во всех деталях."}, 
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}
                ]}
            ], "temperature": 0.4
        }
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
                data = await resp.json()
                base_reply = data["choices"][0]["message"]["content"] if "choices" in data else "Сбой оптического контура."
                return base_reply + await reminder_task

    history = get_history(user_id, limit=6)
    messages = [{"role": "system", "content": full_system_instruction}]
    for h in history: messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": user_input})
    
    payload = {"model": GROQ_MODEL_TEXT, "messages": messages, "temperature": 0.6}
    async with aiohttp.ClientSession() as session:
        async with session.post(GROQ_CHAT_URL, json=payload, headers=headers) as resp:
            data = await resp.json()
            base_reply = data["choices"][0]["message"]["content"] if "choices" in data else "Деградация нейросетевой матрицы."
            return base_reply + await reminder_task

async def respond_fast(message: Message, text_reply: str, user_id: int, custom_text_log: str = None):
    voice_path = f"reply_{user_id}.ogg"
    display_text = custom_text_log if custom_text_log else text_reply
    
    if "[GENERATE_DOC_DOCX]" in text_reply:
        clean_content = text_reply.replace("[GENERATE_DOC_DOCX]", "").strip()
        if len(clean_content) > 30:
            fn = f"report_{user_id}.docx"
            create_docx_report("Аналитический документ J.A.R.V.I.S.", clean_content, fn)
            await message.answer_document(FSInputFile(fn), caption="Сэр, ваш документ Word сформирован.")
            if os.path.exists(fn): os.remove(fn)
        display_text = display_text.replace("[GENERATE_DOC_DOCX]", "")
        text_reply = text_reply.replace("[GENERATE_DOC_DOCX]", "")
        
    elif "[GENERATE_DOC_XLSX]" in text_reply:
        lines = text_reply.replace("[GENERATE_DOC_XLSX]", "").strip().split("\n")
        table = [line.split("|") for line in lines if line]
        if len(table) > 0:
            fn = f"matrix_{user_id}.xlsx"
            create_xlsx_report(table, fn)
            await message.answer_document(FSInputFile(fn), caption="Сэр, таблица Excel подготовлена.")
            if os.path.exists(fn): os.remove(fn)
        display_text = display_text.replace("[GENERATE_DOC_XLSX]", "")
        text_reply = text_reply.replace("[GENERATE_DOC_XLSX]", "")

    try:
        await message.answer(escape_markdown(display_text), parse_mode="MarkdownV2")
    except Exception:
        await message.answer(display_text, parse_mode=None)
        
    async def generate_and_send_voice():
        await edge_tts.Communicate(text_reply.replace("*","").replace("_",""), JARVIS_VOICE).save(voice_path)
        if os.path.exists(voice_path) and os.path.getsize(voice_path) > 0:
            try: await message.answer_voice(voice=FSInputFile(voice_path))
            except: pass
        if os.path.exists(voice_path): os.remove(voice_path)
    asyncio.create_task(generate_and_send_voice())

# ─── МУЛЬТИФОРМАТНЫЕ ОБРАБОТЧИКИ ВХОДЯЩИХ ДАННЫХ ──────────────────────────────

async def process_and_transcribe_audio(file_id: int, user_id: int) -> str:
    """Единый узел дешифрации медиа-звука через Groq Whisper"""
    local_path = f"audio_{user_id}.ogg"
    try:
        file = await bot.get_file(file_id)
        await bot.download_file(file.file_path, local_path)
        
        headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
        data = aiohttp.FormData()
        data.add_field("file", open(local_path, "rb"), filename="voice.ogg")
        data.add_field("model", "whisper-large-v3")
        data.add_field("language", "ru")
        
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_AUDIO_URL, data=data, headers=headers) as resp:
                res = await resp.json()
                return res.get("text", "")
    except Exception as e:
        logger.error(f"Audio processing error: {e}")
        return ""
    finally:
        if os.path.exists(local_path): os.remove(local_path)

# --- ИСПРАВЛЕННЫЙ ВСЕЯДНЫЙ ОБРАБОТЧИК ---

@dp.message(F.document | F.photo | F.audio | F.voice)
async def handle_any_file(message: Message):
    user_id = message.from_user.id
    init_db()
    
    # 1. Если это фото
    if message.photo:
        await message.answer("📸 Получено фото. Включаю визуальный анализ...")
        photo = message.photo[-1]
        local_path = f"img_{user_id}.jpg"
        await bot.download(photo, destination=local_path)
        with open(local_path, "rb") as img_file:
            img_b64 = base64.b64encode(img_file.read()).decode('utf-8')
        os.remove(local_path)
        reply = await process_jarvis_thought(user_id, "Проанализируй это изображение", image_b64=img_b64)
        await respond_fast(message, reply, user_id)
        return

    # 2. Если это аудио или голос
    if message.audio or message.voice:
        await message.answer("🎧 Анализирую аудиопоток...")
        file_id = message.audio.file_id if message.audio else message.voice.file_id
        text_input = await process_and_transcribe_audio(file_id, user_id)
        if text_input:
            reply = await process_jarvis_thought(user_id, f"Анализ аудио: {text_input}")
            await respond_fast(message, reply, user_id, custom_text_log=f"📝 *Текст:* {text_input}\n\n{reply}")
        return

    # 3. Если это любой другой документ
    if message.document:
        file_name = message.document.file_name
        file_ext = file_name.split(".")[-1].lower()
        local_tmp_path = f"tmp_{user_id}_{file_name}"
        await bot.download_file((await bot.get_file(message.document.file_id)).file_path, local_tmp_path)
        
        content = ""
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
            else:
                # Если формат неизвестен, пробуем прочитать как текст
                with open(local_tmp_path, "r", encoding="utf-8", errors="ignore") as f: content = f.read(5000)
            
            if content:
                add_to_knowledge_db(user_id, file_name, content)
                await message.answer(f"✅ Файл `{file_name}` принят и проиндексирован.")
            else:
                await message.answer("⚠️ Не удалось извлечь текст из этого файла.")
        finally:
            if os.path.exists(local_tmp_path): os.remove(local_tmp_path)

@dp.message(F.voice)
async def handle_telegram_voice(message: Message):
    user_id = message.from_user.id
    init_db()
    text_input = await process_and_transcribe_audio(message.voice.file_id, user_id)
    if not text_input: return
    
    save_message(user_id, "user", text_input)
    reply = await process_jarvis_thought(user_id, text_input)
    save_message(user_id, "assistant", reply)
    await respond_fast(message, reply, user_id, custom_text_log=f"🗣 *Голос:* {text_input}\n\n{reply}")

@dp.message(F.audio)
async def handle_incoming_audio_file(message: Message):
    """Сенсор обработки полноценных аудиозаписей (MP3, WAV лекции)"""
    user_id = message.from_user.id
    init_db()
    await message.answer("📥 *Протокол Аудио:* Получен файл записи лекции/аудио. Расшифровываю звуковую дорожку...")
    
    text_input = await process_and_transcribe_audio(message.audio.file_id, user_id)
    if not text_input:
        await message.answer("❌ Сэр, не удалось декодировать аудиопоток. Файл поврежден или имеет неверный кодек.")
        return
        
    save_message(user_id, "user", f"[Расшифровка аудио {message.audio.file_name}]: {text_input}")
    reply = await process_jarvis_thought(user_id, f"Проанализируй текст этой аудиозаписи: {text_input}")
    save_message(user_id, "assistant", reply)
    await respond_fast(message, reply, user_id, custom_text_log=f"📝 *Текст аудиозаписи:* {text_input}\n\n{reply}")

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    init_db()
    save_message(user_id, "user", message.text)
    reply = await process_jarvis_thought(user_id, message.text)
    save_message(user_id, "assistant", reply)
    await respond_fast(message, reply, user_id)

async def main():
    init_db()
    scheduler.start()
    logger.info("Ультимативный мультиформатный Джарвис активен.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
