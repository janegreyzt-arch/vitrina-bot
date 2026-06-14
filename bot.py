import asyncio
import logging
from datetime import datetime, timedelta
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile
import aiosqlite
import re
import os
from dotenv import load_dotenv
from rapidfuzz import process, fuzz
from openpyxl import Workbook

load_dotenv()
TOKEN = os.getenv("BOT_TOKEN")

bot = Bot(token=TOKEN)
dp = Dispatcher()

# ====================== НАСТРОЙКИ ======================
POINTS = ["МН", "СМ", "Д1", "СОК", "ПТ", "ПК", "МСТИЛЬ"]
TIMES = ["10:00", "16:00", "20:00"]
ADMIN_CODE = "1506"
ADMIN_IDS = []   # список активных админов

# ====================== БАЗА ДАННЫХ ======================
async def init_db():
    async with aiosqlite.connect("vitrina_bot.db") as db:
        await db.execute("""CREATE TABLE IF NOT EXISTS vitrina_reports (
            id INTEGER PRIMARY KEY, date TEXT, point TEXT, time_slot TEXT,
            user_id INTEGER, username TEXT, message_time TEXT, status TEXT
        )""")
        await db.execute("""CREATE TABLE IF NOT EXISTS movements (
            id INTEGER PRIMARY KEY, date TEXT, point TEXT, user_id INTEGER,
            username TEXT, action TEXT, time TEXT
        )""")
        await db.commit()

def get_today():
    return datetime.now().strftime("%Y-%m-%d")

def normalize_text(text: str) -> str:
    return re.sub(r'[^а-яА-Я0-9:]', '', text.upper())

VITRINA_KEYWORDS = ["ВИТРИНА", "ВИТРИН", "ВИТР"]
LEAVE_KEYWORDS = ["ВЫШЕЛ", "ВЫШЛА", "ОТОШЕЛ", "ОТОШЛА", "ПОЕХАЛ", "ПОЕХАЛА", "УШЕЛ", "УШЛА", "ВЫЕХАЛ"]
RETURN_KEYWORDS = ["ВЕРНУЛСЯ", "ВЕРНУЛАСЬ", "ПРИЕХАЛ", "ПРИЕХАЛА", "ПРИБЫЛ", "ПРИБЫЛА"]

# ====================== ОБРАБОТКА ======================
@dp.message(F.chat.type.in_({"group", "supergroup"}))
async def handle_message(message: Message):
    if not message.text: return
    text = message.text.strip()
    norm_text = normalize_text(text)
    user_id = message.from_user.id
    username = message.from_user.username or message.from_user.full_name
    msg_time = datetime.now()

    if any(kw in norm_text for kw in VITRINA_KEYWORDS):
        point_match = process.extractOne(norm_text, POINTS, scorer=fuzz.partial_ratio)
        if point_match and point_match[1] >= 70:
            point = point_match[0]
            time_match = re.search(r'(\d{1,2})[:.]?(\d{2})?', text)
            if time_match:
                h = int(time_match.group(1))
                m = int(time_match.group(2)) if time_match.group(2) else 0
                time_slot = f"{h:02d}:{m:02d}"
                if time_slot in TIMES:
                    await save_vitrina(point, time_slot, user_id, username, msg_time)
                    await bot.send_message(user_id, f"Витрина {point} на {time_slot} принята")
                    return

    action = None
    if any(kw in norm_text for kw in LEAVE_KEYWORDS):
        action = "leave"
    elif any(kw in norm_text for kw in RETURN_KEYWORDS):
        action = "return"

    if action:
        point_match = process.extractOne(norm_text, POINTS, scorer=fuzz.partial_ratio)
        if point_match and point_match[1] >= 65:
            point = point_match[0]
            await save_movement(point, action, user_id, username, msg_time)
            status = "вышел" if action == "leave" else "вернулся"
            await bot.send_message(user_id, f"{point} — {status}")
            return

# ====================== СОХРАНЕНИЕ ======================
async def save_vitrina(point, time_slot, user_id, username, msg_time):
    today = get_today()
    expected = datetime.strptime(time_slot, "%H:%M").time()
    delta = timedelta(minutes=30)
    lower_time = (datetime.combine(datetime.today(), expected) - delta).time()
    upper_time = (datetime.combine(datetime.today(), expected) + delta).time()
    status = "on_time" if lower_time <= msg_time.time() <= upper_time else "late"

    async with aiosqlite.connect("vitrina_bot.db") as db:
        await db.execute("INSERT INTO vitrina_reports VALUES (NULL,?,?,?,?,?,?,?)",
            (today, point, time_slot, user_id, username, msg_time.isoformat(), status))
        await db.commit()

async def save_movement(point, action, user_id, username, msg_time):
    today = get_today()
    async with aiosqlite.connect("vitrina_bot.db") as db:
        await db.execute("INSERT INTO movements VALUES (NULL,?,?,?,?,?,?)",
            (today, point, user_id, username, action, msg_time.strftime("%H:%M")))
        await db.commit()

# ====================== АДМИН КОМАНДЫ ======================
@dp.message(Command("admin"))
async def cmd_admin(message: Message):
    await bot.send_message(message.from_user.id, "Введите код доступа:")

@dp.message(Command("admins"))
async def cmd_admins(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if not ADMIN_IDS:
        await bot.send_message(message.from_user.id, "Список админов пуст.")
        return
    
    text = "👥 Список админов:\n\n"
    for uid in ADMIN_IDS:
        text += f"• ID: {uid}\n"
    await bot.send_message(message.from_user.id, text)

@dp.message()
async def handle_admin_code(message: Message):
    global ADMIN_IDS
    if message.text and message.text.strip() == ADMIN_CODE:
        user_id = message.from_user.id
        if user_id not in ADMIN_IDS:
            ADMIN_IDS.append(user_id)
            await bot.send_message(user_id, 
                "✅ Доступ разрешён!\n\n"
                "Команды:\n"
                "/report_vitrina [дата]\n"
                "/report_movement [дата]\n"
                "/admins — посмотреть список админов")
        else:
            await bot.send_message(user_id, "Вы уже имеете доступ.")

# ====================== ОТЧЕТЫ (без изменений) ======================
async def generate_vitrina_excel(date: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Витрины"
    ws.append(["Точка", "Запланировано", "Статус", "Сотрудник", "Время отправки"])

    async with aiosqlite.connect("vitrina_bot.db") as db:
        async with db.execute("SELECT point, time_slot, status, username, message_time FROM vitrina_reports WHERE date = ? ORDER BY point, time_slot", (date,)) as cursor:
            rows = await cursor.fetchall()
            for row in rows:
                status_text = "Вовремя" if row[2] == "on_time" else "Опоздание"
                ws.append([row[0], row[1], status_text, row[3], row[4][:16]])

    filename = f"Витрины_{date}.xlsx"
    wb.save(filename)
    return filename

async def generate_movement_excel(date: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Выходы и возвраты"
    ws.append(["Точка", "Действие", "Время", "Сотрудник"])

    async with aiosqlite.connect("vitrina_bot.db") as db:
        async with db.execute("SELECT point, action, time, username FROM movements WHERE date = ? ORDER BY time", (date,)) as cursor:
            rows = await cursor.fetchall()
            for row in rows:
                action_text = "Вышел" if row[1] == "leave" else "Вернулся"
                ws.append([row[0], action_text, row[2], row[3]])

    filename = f"Движения_{date}.xlsx"
    wb.save(filename)
    return filename

@dp.message(Command("report_vitrina"))
async def report_vitrina(message: Message):
    if message.from_user.id not in ADMIN_IDS: return
    date = message.text.split()[-1] if len(message.text.split()) > 1 else get_today()
    if not re.match(r"\d{4}-\d{2}-\d{2}", date):
        date = get_today()
    filename = await generate_vitrina_excel(date)
    await bot.send_document(message.from_user.id, FSInputFile(filename), caption=f"Отчет по витринам за {date}")

@dp.message(Command("report_movement"))
async def report_movement(message: Message):
    if message.from_user.id not in ADMIN_IDS: return
    date = message.text.split()[-1] if len(message.text.split()) > 1 else get_today()
    if not re.match(r"\d{4}-\d{2}-\d{2}", date):
        date = get_today()
    filename = await generate_movement_excel(date)
    await bot.send_document(message.from_user.id, FSInputFile(filename), caption=f"Отчет по движениям за {date}")

# ====================== ЗАПУСК ======================
async def main():
    await init_db()
    logging.basicConfig(level=logging.INFO)
    print("✅ Бот успешно запущен на Bothost!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())