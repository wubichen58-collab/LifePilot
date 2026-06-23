import asyncio
import logging
import os
import sqlite3
import json
from datetime import datetime, time
import aiohttp

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import Message
from apscheduler.schedulers.asyncio import AsyncIOScheduler

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key={GEMINI_API_KEY}"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
scheduler = AsyncIOScheduler()

# ─── DATABASE ────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS profile (
            user_id INTEGER PRIMARY KEY,
            name TEXT DEFAULT 'Амаль',
            goals TEXT DEFAULT '["HSK 4", "блог", "отношения на расстоянии"]',
            energy TEXT DEFAULT 'средняя',
            mood TEXT DEFAULT 'нормально',
            money INTEGER DEFAULT 0,
            yesterday TEXT DEFAULT '',
            schedule TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS chat_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            role TEXT,
            content TEXT,
            created_at TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS finances (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            amount INTEGER,
            description TEXT,
            created_at TEXT
        )
    """)
    conn.commit()
    conn.close()

def get_profile(user_id: int) -> dict:
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    c.execute("SELECT * FROM profile WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    if not row:
        create_profile(user_id)
        return get_profile(user_id)
    keys = ["user_id","name","goals","energy","mood","money","yesterday","schedule","notes","updated_at"]
    profile = dict(zip(keys, row))
    profile["goals"] = json.loads(profile["goals"])
    return profile

def create_profile(user_id: int):
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    c.execute("INSERT OR IGNORE INTO profile (user_id) VALUES (?)", (user_id,))
    conn.commit()
    conn.close()

def update_profile(user_id: int, **kwargs):
    if "goals" in kwargs and isinstance(kwargs["goals"], list):
        kwargs["goals"] = json.dumps(kwargs["goals"], ensure_ascii=False)
    kwargs["updated_at"] = datetime.now().isoformat()
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    for key, value in kwargs.items():
        c.execute(f"UPDATE profile SET {key} = ? WHERE user_id = ?", (value, user_id))
    conn.commit()
    conn.close()

def add_finance(user_id: int, amount: int, description: str):
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    c.execute(
        "INSERT INTO finances (user_id, amount, description, created_at) VALUES (?, ?, ?, ?)",
        (user_id, amount, description, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()
    profile = get_profile(user_id)
    new_balance = profile["money"] + amount
    update_profile(user_id, money=new_balance)
    return new_balance

def get_recent_finances(user_id: int, limit: int = 5) -> list:
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    c.execute(
        "SELECT amount, description, created_at FROM finances WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, limit)
    )
    rows = c.fetchall()
    conn.close()
    return rows

def save_message(user_id: int, role: str, content: str):
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    c.execute(
        "INSERT INTO chat_history (user_id, role, content, created_at) VALUES (?, ?, ?, ?)",
        (user_id, role, content, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()

def get_history(user_id: int, limit: int = 10) -> list:
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    c.execute(
        "SELECT role, content FROM chat_history WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, limit)
    )
    rows = c.fetchall()
    conn.close()
    return [{"role": r, "parts": [{"text": c}]} for r, c in reversed(rows)]

# ─── GEMINI ──────────────────────────────────────────────────────────────────

async def ask_gemini(system_prompt: str, user_message: str, history: list = None) -> str:
    contents = []
    if history:
        contents.extend(history)
    contents.append({"role": "user", "parts": [{"text": user_message}]})

    payload = {
        "system_instruction": {"parts": [{"text": system_prompt}]},
        "contents": contents,
        "generationConfig": {"maxOutputTokens": 1024, "temperature": 0.8}
    }

    async with aiohttp.ClientSession() as session:
        async with session.post(GEMINI_URL, json=payload) as resp:
            data = await resp.json()
            try:
                return data["candidates"][0]["content"]["parts"][0]["text"]
            except Exception:
                logger.error(f"Gemini error: {data}")
                return "Ошибка при обращении к Gemini. Попробуй ещё раз."

def build_system_prompt(profile: dict) -> str:
    now = datetime.now()
    return f"""Ты — LifePilot, личный AI-компаньон пользователя по имени {profile['name']}.

Контекст о пользователе:
- Цели: {', '.join(profile['goals'])}
- Энергия сейчас: {profile['energy']}
- Настроение: {profile['mood']}
- Баланс: {profile['money']} юаней
- Вчера делал: {profile['yesterday'] or 'не указано'}
- Расписание: {profile['schedule'] or 'не указано'}
- Заметки: {profile['notes'] or 'нет'}
- Сейчас: {now.strftime('%A, %d %B %Y, %H:%M')}

Твои принципы:
1. Ты знаешь пользователя лично — отвечай конкретно, не обобщённо
2. Не давай списки задач — давай конкретный совет под текущий контекст
3. Будь честным, иногда провокационным, всегда на стороне пользователя
4. Отвечай на русском языке
5. Короткие ответы — максимум 3-4 абзаца, без лишней воды"""

# ─── MORNING BRIEFING ────────────────────────────────────────────────────────

async def send_morning_briefing(user_id: int):
    profile = get_profile(user_id)
    now = datetime.now()

    prompt = build_system_prompt(profile)
    message = f"""Сделай утренний брифинг на сегодня, {now.strftime('%d %B')}.

Структура:
1. Одно предложение про день (тон, настрой)
2. Конкретный план на день (не список — живой текст)
3. Одна вещь которую НЕ стоит делать сегодня
4. Мотивирующая мысль под контекст пользователя

Учти энергию ({profile['energy']}), настроение ({profile['mood']}), и что было вчера."""

    response = await ask_gemini(prompt, message)
    await bot.send_message(user_id, f"🌅 *Доброе утро, {profile['name']}!*\n\n{response}", parse_mode="Markdown")

# ─── COMMANDS ────────────────────────────────────────────────────────────────

@dp.message(Command("start"))
async def cmd_start(message: Message):
    user_id = message.from_user.id
    create_profile(user_id)
    text = """👋 Привет! Я *LifePilot* — твой личный AI-компаньон.

Я знаю твой контекст и помогаю принимать решения: что делать сегодня, как не залипнуть, как поговорить с кем-то важным, куда уходят деньги.

*Команды:*
/profile — посмотреть свой профиль
/energy — обновить уровень энергии
/mood — обновить настроение
/money — финансы
/yesterday — что делал вчера
/schedule — расписание на сегодня
/morning — утренний брифинг прямо сейчас
/blog — идеи для блога
/social — помощь в разговоре или ситуации
/goals — обновить цели

Или просто напиши мне что угодно — я отвечу с учётом твоего контекста."""
    await message.answer(text, parse_mode="Markdown")

@dp.message(Command("profile"))
async def cmd_profile(message: Message):
    profile = get_profile(message.from_user.id)
    finances = get_recent_finances(message.from_user.id)
    fin_text = "\n".join([f"  {'➕' if a > 0 else '➖'} {abs(a)}¥ — {d}" for a, d, _ in finances]) or "  нет записей"
    text = f"""📋 *Твой профиль*

👤 Имя: {profile['name']}
🎯 Цели: {', '.join(profile['goals'])}
⚡ Энергия: {profile['energy']}
😊 Настроение: {profile['mood']}
💰 Баланс: {profile['money']} юаней
📅 Расписание: {profile['schedule'] or 'не указано'}
📝 Вчера: {profile['yesterday'] or 'не указано'}

*Последние транзакции:*
{fin_text}"""
    await message.answer(text, parse_mode="Markdown")

@dp.message(Command("energy"))
async def cmd_energy(message: Message):
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.answer("Укажи уровень энергии:\n/energy высокая\n/energy средняя\n/energy низкая")
        return
    level = args[1].strip().lower()
    update_profile(message.from_user.id, energy=level)
    await message.answer(f"⚡ Энергия обновлена: *{level}*", parse_mode="Markdown")

@dp.message(Command("mood"))
async def cmd_mood(message: Message):
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.answer("Напиши своё настроение:\n/mood отлично\n/mood тревожно\n/mood устал\n/mood мотивирован")
        return
    mood = args[1].strip()
    update_profile(message.from_user.id, mood=mood)
    await message.answer(f"😊 Настроение обновлено: *{mood}*", parse_mode="Markdown")

@dp.message(Command("yesterday"))
async def cmd_yesterday(message: Message):
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.answer("Напиши что делал вчера:\n/yesterday занимался HSK, смотрел сериал, гулял")
        return
    yesterday = args[1].strip()
    update_profile(message.from_user.id, yesterday=yesterday)
    await message.answer(f"📝 Записал: *{yesterday}*", parse_mode="Markdown")

@dp.message(Command("schedule"))
async def cmd_schedule(message: Message):
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.answer("Напиши расписание на сегодня:\n/schedule пары с 9 до 12, свободен после 14")
        return
    schedule = args[1].strip()
    update_profile(message.from_user.id, schedule=schedule)
    await message.answer(f"📅 Расписание обновлено: *{schedule}*", parse_mode="Markdown")

@dp.message(Command("morning"))
async def cmd_morning(message: Message):
    await message.answer("🌅 Готовлю брифинг...")
    await send_morning_briefing(message.from_user.id)

@dp.message(Command("money"))
async def cmd_money(message: Message):
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        profile = get_profile(message.from_user.id)
        await message.answer(f"💰 Текущий баланс: *{profile['money']} юаней*\n\nДобавить доход: /money +2000 стипендия\nЗаписать расход: /money -150 еда", parse_mode="Markdown")
        return

    raw = args[1].strip()
    parts = raw.split(maxsplit=1)
    try:
        amount = int(parts[0].replace("+", ""))
        description = parts[1] if len(parts) > 1 else "без описания"
        new_balance = add_finance(message.from_user.id, amount, description)

        profile = get_profile(message.from_user.id)
        prompt = build_system_prompt(profile)
        comment = await ask_gemini(prompt, f"Пользователь {'получил' if amount > 0 else 'потратил'} {abs(amount)} юаней ({description}). Баланс теперь {new_balance}. Дай короткий (1 предложение) умный комментарий про эту трату в контексте его целей.")

        sign = "➕" if amount > 0 else "➖"
        await message.answer(f"{sign} *{abs(amount)}¥* — {description}\n💰 Баланс: *{new_balance}¥*\n\n💡 {comment}", parse_mode="Markdown")
    except (ValueError, IndexError):
        await message.answer("Формат: /money +2000 стипендия или /money -150 еда")

@dp.message(Command("blog"))
async def cmd_blog(message: Message):
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.answer("Напиши о чём был твой день или событие:\n/blog сегодня ездили на картинг с друзьями")
        return
    event = args[1].strip()
    profile = get_profile(message.from_user.id)
    prompt = build_system_prompt(profile)
    response = await ask_gemini(prompt, f"""Пользователь хочет сделать контент про: {event}

Предложи:
1. 3 идеи для поста (разные форматы)
2. 2 варианта заголовка для Reels
3. Первое предложение для подписи (цепляющее)

Учти его стиль жизни и цели.""")
    await message.answer(f"🎥 *Идеи для контента*\n\n{response}", parse_mode="Markdown")

@dp.message(Command("social"))
async def cmd_social(message: Message):
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.answer("Опиши ситуацию:\n/social поссорился с другом из-за планов\n/social хочу поговорить с родителями о переезде")
        return
    situation = args[1].strip()
    profile = get_profile(message.from_user.id)
    prompt = build_system_prompt(profile)
    response = await ask_gemini(prompt, f"""Ситуация: {situation}

Помоги разобраться:
1. Как ты видишь эту ситуацию (честно)
2. Два варианта как можно действовать
3. Что сказать — конкретная фраза для начала разговора

Будь прямым, не давай банальных советов.""")
    await message.answer(f"❤️ *Социальный навигатор*\n\n{response}", parse_mode="Markdown")

@dp.message(Command("goals"))
async def cmd_goals(message: Message):
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        profile = get_profile(message.from_user.id)
        await message.answer(f"🎯 Текущие цели: {', '.join(profile['goals'])}\n\nОбновить:\n/goals HSK 4, блог 1000 подписчиков, поездка в Казахстан")
        return
    goals_raw = args[1].strip()
    goals = [g.strip() for g in goals_raw.split(",")]
    update_profile(message.from_user.id, goals=goals)
    await message.answer(f"🎯 Цели обновлены:\n" + "\n".join([f"• {g}" for g in goals]))

# ─── FREE CHAT ───────────────────────────────────────────────────────────────

@dp.message(F.text)
async def free_chat(message: Message):
    user_id = message.from_user.id
    profile = get_profile(user_id)
    history = get_history(user_id, limit=8)
    prompt = build_system_prompt(profile)

    save_message(user_id, "user", message.text)
    response = await ask_gemini(prompt, message.text, history)
    save_message(user_id, "model", response)

    await message.answer(response)

# ─── SCHEDULER ───────────────────────────────────────────────────────────────

async def schedule_morning_briefings():
    conn = sqlite3.connect("lifepilot.db")
    c = conn.cursor()
    c.execute("SELECT user_id FROM profile")
    users = c.fetchall()
    conn.close()
    for (user_id,) in users:
        try:
            await send_morning_briefing(user_id)
        except Exception as e:
            logger.error(f"Morning briefing error for {user_id}: {e}")

# ─── MAIN ────────────────────────────────────────────────────────────────────

async def main():
    init_db()
    scheduler.add_job(schedule_morning_briefings, "cron", hour=8, minute=0, timezone="Asia/Shanghai")
    scheduler.start()
    logger.info("LifePilot started!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
