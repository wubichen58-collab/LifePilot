import asyncio
import logging
import os
import sqlite3
import json
import base64
import re
from datetime import datetime, timedelta
import aiohttp

# Библиотеки дешифрации
import docx
import openpyxl
import pypdf
import edge_tts

from aiogram import Bot, Dispatcher, types, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import Command
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from textblob import TextBlob
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.linear_model import LogisticRegression

# ─── 1. ИНИЦИАЛИЗАЦИЯ И НАСТРОЙКИ ─────────────────────────────────────────────
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")

GROQ_CHAT_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_AUDIO_URL = "https://api.groq.com/openai/v1/audio/transcriptions"
GROQ_MODEL_TEXT = "llama-3.3-70b-versatile"
GROQ_MODEL_VISION = "llama-3.2-11b-vision-preview"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
scheduler = AsyncIOScheduler()

# ─── 2. МАШИННОЕ ОБУЧЕНИЕ (ЭМОЦИОНАЛЬНЫЙ ФОН) ─────────────────────────────────
TRAIN_TEXTS = ["надо срочно сделать отчет", "дедлайн завтра", "исправь ошибку", "привет как дела", "расскажи шутку", "что делаешь", "я устал", "все надоело", "какой-то бред", "ничего не получается"]
TRAIN_LABELS = [0, 0, 0, 1, 1, 1, 2, 2, 2, 2]

vectorizer = CountVectorizer()
X_train = vectorizer.fit_transform(TRAIN_TEXTS)
ml_classifier = LogisticRegression()
ml_classifier.fit(X_train, TRAIN_LABELS)

def predict_context_ml(text: str) -> str:
    try:
        prediction = ml_classifier.predict(vectorizer.transform([text.lower()]))[0]
        if prediction == 0: return "Деловой/Критический"
        if prediction == 2: return "Повышенный стресс/Усталость"
        return "Стандартный/Развлекательный"
    except: return "Не определен"

# ─── 3. БАЗА ДАННЫХ (RAG И ДОЛГАЯ ПАМЯТЬ) ─────────────────────────────────────
def init_db():
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS consciousness (user_id INTEGER PRIMARY KEY, user_name TEXT DEFAULT 'Сэр', shared_interests TEXT DEFAULT '{}', jarvis_opinion_matrix TEXT DEFAULT '{}', money INTEGER DEFAULT 0, system_log TEXT DEFAULT 'Инициализация.')")
    c.execute("CREATE TABLE IF NOT EXISTS chat_history (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, role TEXT, content TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS reminders (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, time TEXT, task TEXT, status TEXT DEFAULT 'pending')")
    c.execute("CREATE TABLE IF NOT EXISTS activity_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, timestamp TEXT, hour INTEGER, category TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS knowledge_base (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, source_name TEXT, content_chunk TEXT)")
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
    """Извлекает последние сообщения для долгой памяти"""
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("SELECT role, content FROM chat_history WHERE user_id = ? ORDER BY id DESC LIMIT ?", (user_id, limit))
    rows = c.fetchall()
    conn.close()
    return [{"role": r[0], "content": r[1]} for r in reversed(rows)]

def add_to_knowledge_db(user_id: int, source_name: str, content: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    for i in range(0, len(content), 800):
        c.execute("INSERT INTO knowledge_base (user_id, source_name, content_chunk) VALUES (?, ?, ?)", (user_id, source_name, content[i:i+1000]))
    conn.commit()
    conn.close()

def query_knowledge_db(user_id: int, query: str, limit: int = 4) -> str:
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    words = [w for w in re.split(r'\W+', query.lower()) if len(w) > 3]
    if not words: conn.close(); return ""
    like_cond = " OR ".join(["content_chunk LIKE ?" for _ in words])
    c.execute(f"SELECT source_name, content_chunk FROM knowledge_base WHERE user_id = ? AND ({like_cond}) LIMIT ?", [user_id] + [f"%{w}%" for w in words] + [limit])
    rows = c.fetchall()
    conn.close()
    return "\n--- ЛОКАЛЬНАЯ БАЗА (RAG) ---\n" + "\n".join([f"[{r[0]}]: {r[1]}" for r in rows]) if rows else ""

def log_user_activity(user_id: int, category: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("INSERT INTO activity_logs (user_id, timestamp, hour, category) VALUES (?, ?, ?, ?)", (user_id, datetime.now().isoformat(), datetime.now().hour, category))
    conn.commit()
    conn.close()

def add_reminder_db(user_id: int, run_time: datetime, task: str):
    conn = sqlite3.connect("jarvis_consciousness.db")
    c = conn.cursor()
    c.execute("INSERT INTO reminders (user_id, time, task) VALUES (?, ?, ?)", (user_id, run_time.isoformat(), task))
    conn.commit()
    conn.close()

# ─── 4. WEB ПОИСК, ДОКУМЕНТЫ И УТИЛИТЫ ────────────────────────────────────────
async def search_web_ddg(query: str) -> str:
    url = f"https://html.duckduckgo.com/html/?q={query}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers, timeout=5) as resp:
                if resp.status != 200: return ""
                html = await resp.text()
                snippets = re.findall(r'<a class="result__snippet".*?>(.*?)</a>', html, re.DOTALL)
                clean_snips = [re.sub(r'<[^>]*>', '', s).strip() for s in snippets[:3]]
                return "\n--- ИНТЕРНЕТ ---\n" + "\n".join(clean_snips) if clean_snips else ""
    except: return ""

def create_docx_report(title: str, content: str, filename: str):
    doc = docx.Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"Дата генерации: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\nСистема: J.A.R.V.I.S.\n")
    doc.add_paragraph(content)
    doc.save(filename)

def create_xlsx_report(table_data: list, filename: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in table_data: ws.append(row)
    wb.save(filename)

def escape_markdown(text: str) -> str:
    return re.sub(r'([_*\[\]()~`>#+\-=|{}.!])', r'\\\1', text)

def get_voice_for_text(text: str) -> str:
    """Анализирует текст и выбирает подходящий голос (RU, EN, ZH)"""
    if re.search(r'[\u4e00-\u9fff]', text):
        return "zh-CN-YunxiNeural" # Мужской китайский голос
    elif re.search(r'[a-zA-Z]', text) and not re.search(r'[\u0400-\u04FF]', text):
        return "en-GB-RyanNeural" # Британский английский голос
    else:
        return "ru-RU-DmitryNeural" # По умолчанию русский

# ─── 5. ПЛАНИРОВЩИК ЗАДАЧ ─────────────────────────────────────────────────────
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
    except Exception as e: logger.error(f"Scheduler error: {e}")

async def check_and_extract_reminders(user_id: int, user_input: str):
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    sys_inst = "Если в тексте есть задача на будущее, верни JSON: {\"has_reminder\": true, \"minutes_delay\": число, \"task\": \"суть\"}. Иначе: {\"has_reminder\": false}"
    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [{"role": "system", "content": sys_inst}, {"role": "user", "content": user_input}],
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

# ─── 6. ЦЕНТРАЛЬНЫЙ МОЗГ (LLM + ДОЛГАЯ ПАМЯТЬ + ПОЛИГЛОТ) ─────────────────────
async def process_jarvis_thought(user_id: int, user_input: str, image_b64: str = None) -> str:
    init_db()
    
    # 1. Достаем ИСТОРИЮ до сохранения текущего сообщения
    history = get_history(user_id, limit=20)
    
    # 2. Сохраняем запрос
    if user_input:
        save_message(user_id, "user", user_input)

    mind = get_mind(user_id)
    ml_context = predict_context_ml(user_input)
    log_user_activity(user_id, ml_context)
    
    rag_context = query_knowledge_db(user_id, user_input)
    web_context = ""
    if any(w in user_input.lower() for w in ["новости", "найти", "сейчас", "курс", "интернет"]):
        web_context = await search_web_ddg(user_input)
        
    system_prompt = f"""Ты — J.A.R.V.I.S., ИИ-ассистент Создателя.
Атмосфера: {ml_context}. Отвечай структурировано, без лишних вступлений.
ВАЖНОЕ ПРАВИЛО ЯЗЫКА: Создатель владеет русским, английским и китайским. Твоя задача — ВСЕГДА отвечать строго на том языке, на котором Создатель написал или произнес последнее сообщение. Не смешивай языки без просьбы.
Всегда запоминай списки и контекст из истории диалога.
Документы: Если просят составить отчет или таблицу, используй [GENERATE_DOC_DOCX] или [GENERATE_DOC_XLSX]."""

    full_system_instruction = f"{system_prompt}\n{rag_context}\n{web_context}"
    
    # 3. Формируем контекст
    messages = [{"role": "system", "content": full_system_instruction}]
    
    for msg in history:
        messages.append({"role": msg["role"], "content": msg["content"]})
        
    if image_b64:
        messages.append({
            "role": "user", 
            "content": [
                {"type": "text", "text": user_input or "Опиши это изображение"}, 
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_b64}"}}
            ]
        })
    else:
        messages.append({"role": "user", "content": user_input})

    reminder_task = asyncio.create_task(check_and_extract_reminders(user_id, user_input))
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
                    # 4. Сохраняем ответ Джарвиса
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
    
    # Генерация документов
    if "[GENERATE_DOC_DOCX]" in text_reply:
        clean_content = text_reply.replace("[GENERATE_DOC_DOCX]", "").strip()
        if len(clean_content) > 10:
            fn = f"report_{user_id}.docx"
            create_docx_report("Документ J.A.R.V.I.S.", clean_content, fn)
            await message.answer_document(FSInputFile(fn), caption="Сэр, ваш документ Word.")
            if os.path.exists(fn): os.remove(fn)
        display_text = display_text.replace("[GENERATE_DOC_DOCX]", "")
        text_reply = text_reply.replace("[GENERATE_DOC_DOCX]", "")
        
    elif "[GENERATE_DOC_XLSX]" in text_reply:
        lines = text_reply.replace("[GENERATE_DOC_XLSX]", "").strip().split("\n")
        table = [line.split("|") for line in lines if line]
        if len(table) > 0:
            fn = f"matrix_{user_id}.xlsx"
            create_xlsx_report(table, fn)
            await message.answer_document(FSInputFile(fn), caption="Сэр, таблица Excel.")
            if os.path.exists(fn): os.remove(fn)
        display_text = display_text.replace("[GENERATE_DOC_XLSX]", "")
        text_reply = text_reply.replace("[GENERATE_DOC_XLSX]", "")

    try:
        await message.answer(escape_markdown(display_text), parse_mode="MarkdownV2")
    except Exception:
        await message.answer(display_text, parse_mode=None)
        
    async def generate_and_send_voice():
        try:
            # Выбор диктора в зависимости от языка ответа
            dynamic_voice = get_voice_for_text(text_reply)
            await edge_tts.Communicate(text_reply.replace("*","").replace("_",""), dynamic_voice).save(voice_path)
            if os.path.exists(voice_path) and os.path.getsize(voice_path) > 0:
                await message.answer_voice(voice=FSInputFile(voice_path))
                os.remove(voice_path)
        except Exception as e: logger.error(f"Voice generation failed: {e}")
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
        # Убрана привязка к русскому языку — теперь Whisper автоматически распознает RU/EN/ZH
        
        async with aiohttp.ClientSession() as session:
            async with session.post(GROQ_AUDIO_URL, data=data, headers=headers) as resp:
                res = await resp.json()
                return res.get("text", "")
    except Exception as e:
        logger.error(f"Audio processing error: {e}")
        return ""
    finally:
        if os.path.exists(local_path): os.remove(local_path)

# ─── 7. НЕЗАВИСИМЫЕ ОБРАБОТЧИКИ ───────────────────────────────────────────────

@dp.message(F.photo)
async def handle_photo(message: Message):
    user_id = message.from_user.id
    init_db()
    await message.answer("📸 Получено фото. Включаю визуальный анализ...")
    
    photo = message.photo[-1]
    local_path = f"img_{user_id}.jpg"
    await bot.download(photo, destination=local_path)
    
    with open(local_path, "rb") as img_file:
        img_b64 = base64.b64encode(img_file.read()).decode('utf-8')
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
            with open(local_tmp_path, "r", encoding="utf-8", errors="ignore") as f: content = f.read(5000)
        
        if content:
            add_to_knowledge_db(user_id, file_name, content)
            await message.answer(f"✅ Файл `{file_name}` успешно загружен в базу знаний (RAG).")
        else:
            await message.answer("⚠️ Не удалось извлечь текст из файла.")
    except Exception as e:
        logger.error(f"Document parse error: {e}")
        await message.answer(f"💥 Ошибка при анализе файла `{file_name}`.")
    finally:
        if os.path.exists(local_tmp_path): os.remove(local_tmp_path)

@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    init_db()
    reply = await process_jarvis_thought(user_id, message.text)
    await respond_fast(message, reply, user_id)

# ─── 8. ЗАПУСК ────────────────────────────────────────────────────────────────
async def main():
    init_db()
    scheduler.start()
    logger.info("Ультимативный мультиформатный мультиязычный Джарвис активен.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
