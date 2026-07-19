import requests
import time
import json
import sqlite3
import os
import re
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font

# ===== ТОКЕН ИЗ ПЕРЕМЕННОЙ ОКРУЖЕНИЯ =====
TG_TOKEN = os.environ.get("TG_TOKEN")
if not TG_TOKEN:
    print("❌ Ошибка: не установлен TG_TOKEN")
    exit()

ADMIN_ID = 7461823442
print("🚀 Запуск финансового помощника...")

# ===== ПРОВЕРКА TELEGRAM =====
url = f"https://api.telegram.org/bot{TG_TOKEN}/getMe"
resp = requests.get(url)
if not resp.json().get("ok"):
    print("❌ Ошибка: неверный Telegram-токен")
    exit()

print("✅ Telegram подключён")
requests.get(f"https://api.telegram.org/bot{TG_TOKEN}/deleteWebhook")
print("🔄 Webhook сброшен")

# ===== БАЗА ДАННЫХ =====
DB_PATH = "finance.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        user_id INTEGER PRIMARY KEY,
        premium INTEGER DEFAULT 0,
        premium_until TEXT,
        first_visit TEXT,
        last_activity TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        type TEXT,
        category TEXT,
        amount REAL,
        description TEXT,
        date TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS budgets (
        user_id INTEGER,
        category TEXT,
        limit_amount REAL,
        PRIMARY KEY (user_id, category)
    )''')
    conn.commit()
    conn.close()

init_db()

# ===== КАТЕГОРИИ =====
CATEGORIES = {
    "income": {
        "💰 Зарплата": "основная работа",
        "💼 Фриланс": "проект, заказ",
        "🎁 Подарок": "деньги на день рождения",
        "📈 Инвестиции": "дивиденды",
        "🏦 Проценты": "банковские проценты",
        "🔄 Возврат": "вернули долг",
        "💰 Другое": "прочий доход"
    },
    "expense": {
        "🍔 Еда": "продукты, кафе",
        "🚇 Транспорт": "такси, бензин",
        "🏠 ЖКХ": "квартплата, свет",
        "🛍️ Покупки": "техника, подарки",
        "💊 Здоровье": "лекарства, врачи",
        "🎉 Развлечения": "кино, игры",
        "📱 Связь": "интернет, телефон",
        "👕 Одежда": "обувь, вещи",
        "📚 Образование": "курсы, книги",
        "💰 Другое": "прочие расходы"
    }
}

# ===== ФУНКЦИЯ ФОРМАТИРОВАНИЯ ЧИСЕЛ =====
def format_amount(amount):
    return f"{amount:,.0f}".replace(",", " ")

# ===== ФУНКЦИЯ ИЗВЛЕЧЕНИЯ ЧИСЛА ИЗ СТРОКИ =====
def extract_amount_and_desc(text):
    match = re.search(r'([\d,\.]+)', text)
    if not match:
        return None, None
    num_str = match.group(1).replace(',', '.')
    try:
        amount = float(num_str)
    except ValueError:
        return None, None
    desc = text[match.end():].strip()
    if desc and desc[0] in ',. ':
        desc = desc[1:].strip()
    return amount, desc

# ===== ФУНКЦИИ БАЗЫ =====
def get_user(user_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT premium, premium_until FROM users WHERE user_id = ?', (user_id,))
    result = c.fetchone()
    conn.close()
    return {"premium": result[0], "premium_until": result[1]} if result else None

def create_user(user_id):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('INSERT OR IGNORE INTO users (user_id, premium, premium_until, first_visit, last_activity) VALUES (?, 0, NULL, ?, ?)',
              (user_id, now, now))
    conn.commit()
    conn.close()

def update_activity(user_id):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('UPDATE users SET last_activity = ? WHERE user_id = ?', (now, user_id))
    conn.commit()
    conn.close()

def add_transaction(user_id, trans_type, category, amount, description=""):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute('INSERT INTO transactions (user_id, type, category, amount, description, date) VALUES (?, ?, ?, ?, ?, ?)',
              (user_id, trans_type, category, amount, description, date))
    conn.commit()
    conn.close()
    update_activity(user_id)

def get_balance(user_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = "income"', (user_id,))
    income = c.fetchone()[0] or 0
    c.execute('SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = "expense"', (user_id,))
    expense = c.fetchone()[0] or 0
    conn.close()
    return income - expense, income, expense

def get_monthly_expense(user_id, category):
    month_start = datetime.now().replace(day=1).strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = "expense" AND category = ? AND date >= ?',
              (user_id, category, month_start))
    result = c.fetchone()[0] or 0
    conn.close()
    return result

def get_transactions_page(user_id, page, per_page=10):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    offset = page * per_page
    c.execute('''SELECT type, category, amount, description, date 
                 FROM transactions WHERE user_id = ? 
                 ORDER BY date DESC LIMIT ? OFFSET ?''',
              (user_id, per_page, offset))
    result = c.fetchall()
    conn.close()
    return result

def get_total_transactions(user_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT COUNT(*) FROM transactions WHERE user_id = ?', (user_id,))
    result = c.fetchone()[0]
    conn.close()
    return result

def get_report_for_month(user_id, year, month):
    month_start = f"{year}-{month:02d}-01"
    next_month = month + 1 if month < 12 else 1
    next_year = year if month < 12 else year + 1
    month_end = f"{next_year}-{next_month:02d}-01"
    
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT type, category, SUM(amount) FROM transactions WHERE user_id = ? AND date >= ? AND date < ? GROUP BY type, category',
              (user_id, month_start, month_end))
    result = c.fetchall()
    conn.close()
    return result

def get_available_months(user_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT DISTINCT strftime("%Y-%m", date) as month FROM transactions WHERE user_id = ? ORDER BY month DESC',
              (user_id,))
    result = [row[0] for row in c.fetchall()]
    conn.close()
    return result

def get_monthly_income_expense(user_id, year, month):
    month_start = f"{year}-{month:02d}-01"
    next_month = month + 1 if month < 12 else 1
    next_year = year if month < 12 else year + 1
    month_end = f"{next_year}-{next_month:02d}-01"
    
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT type, SUM(amount) FROM transactions WHERE user_id = ? AND date >= ? AND date < ? GROUP BY type',
              (user_id, month_start, month_end))
    result = c.fetchall()
    conn.close()
    
    income = 0
    expense = 0
    for trans_type, amount in result:
        if trans_type == "income":
            income = amount
        else:
            expense = amount
    return income, expense

def get_budget(user_id, category):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT limit_amount FROM budgets WHERE user_id = ? AND category = ?', (user_id, category))
    result = c.fetchone()
    conn.close()
    return result[0] if result else None

def set_budget(user_id, category, amount):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('INSERT OR REPLACE INTO budgets (user_id, category, limit_amount) VALUES (?, ?, ?)',
              (user_id, category, amount))
    conn.commit()
    conn.close()

def delete_budget(user_id, category):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM budgets WHERE user_id = ? AND category = ?', (user_id, category))
    conn.commit()
    conn.close()

def delete_all_data(user_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('DELETE FROM transactions WHERE user_id = ?', (user_id,))
    c.execute('DELETE FROM budgets WHERE user_id = ?', (user_id,))
    conn.commit()
    conn.close()

# ===== ЭКСПОРТ В EXCEL =====
def export_to_excel(chat_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT date, type, category, amount, description FROM transactions WHERE user_id = ? ORDER BY date DESC', (chat_id,))
    data = c.fetchall()
    conn.close()
    
    if not data:
        send_message(chat_id, "📋 Нет операций для экспорта.", main_keyboard(chat_id))
        return
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Операции"
    
    # Дата формирования отчёта
    ws.cell(row=1, column=1, value="Отчёт сформирован:")
    ws.cell(row=1, column=2, value=datetime.now().strftime("%d.%m.%Y %H:%M"))
    ws.cell(row=2, column=1, value="")
    
    # Заголовки
    headers = ["Дата", "Тип", "Категория", "Сумма", "Описание"]
    for col_num, header in enumerate(headers, 1):
        ws.cell(row=3, column=col_num, value=header)
    
    row_num = 4
    total_income = 0
    total_expense = 0
    
    for row in data:
        try:
            date_str = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S").strftime("%d.%m.%Y %H:%M")
        except:
            date_str = row[0]
        trans_type = "Доход" if row[1] == "income" else "Расход"
        amount = row[3]
        formatted_amount = f"{amount:,.0f} ₽".replace(",", " ")
        
        ws.cell(row=row_num, column=1, value=date_str)
        ws.cell(row=row_num, column=2, value=trans_type)
        ws.cell(row=row_num, column=3, value=row[2])
        ws.cell(row=row_num, column=4, value=formatted_amount)
        ws.cell(row=row_num, column=5, value=row[4] or "")
        
        if row[1] == "income":
            total_income += amount
        else:
            total_expense += amount
        
        row_num += 1
    
    row_num += 1
    bold_font = Font(bold=True)
    
    ws.cell(row=row_num, column=3, value="ИТОГО ДОХОДЫ:")
    ws.cell(row=row_num, column=4, value=f"{total_income:,.0f} ₽".replace(",", " "))
    ws.cell(row=row_num, column=3).font = bold_font
    
    row_num += 1
    ws.cell(row=row_num, column=3, value="ИТОГО РАСХОДЫ:")
    ws.cell(row=row_num, column=4, value=f"{total_expense:,.0f} ₽".replace(",", " "))
    ws.cell(row=row_num, column=3).font = bold_font
    
    row_num += 1
    balance = total_income - total_expense
    ws.cell(row=row_num, column=3, value="ДОСТУПНЫЙ БАЛАНС:")
    ws.cell(row=row_num, column=4, value=f"{balance:,.0f} ₽".replace(",", " "))
    ws.cell(row=row_num, column=3).font = bold_font
    
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2
    
    filename = f"отчёт_{datetime.now().strftime('%B_%Y')}.xlsx"
    wb.save(filename)
    
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendDocument"
    files = {'document': open(filename, 'rb')}
    data = {'chat_id': chat_id}
    requests.post(url, files=files, data=data)
    
    os.remove(filename)
    send_message(chat_id, "✅ Файл отправлен!", main_keyboard(chat_id))

# ===== СТАТИСТИКА =====
def get_stats():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    c.execute('SELECT COUNT(*) FROM users')
    total_users = c.fetchone()[0]
    
    week_ago = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
    c.execute('SELECT COUNT(*) FROM users WHERE last_activity >= ?', (week_ago,))
    active_users = c.fetchone()[0]
    
    month_start = datetime.now().replace(day=1).strftime("%Y-%m-%d")
    c.execute('SELECT COUNT(*) FROM transactions WHERE date >= ?', (month_start,))
    total_ops = c.fetchone()[0]
    
    c.execute('SELECT SUM(amount) FROM transactions WHERE type = "income" AND date >= ?', (month_start,))
    total_income = c.fetchone()[0] or 0
    c.execute('SELECT SUM(amount) FROM transactions WHERE type = "expense" AND date >= ?', (month_start,))
    total_expense = c.fetchone()[0] or 0
    
    conn.close()
    return {
        "total_users": total_users,
        "active_users": active_users,
        "total_ops": total_ops,
        "total_income": total_income,
        "total_expense": total_expense
    }

# ===== КЛАВИАТУРЫ =====
def main_keyboard(chat_id):
    keyboard = [
        ["💰 Баланс", "📊 Отчёт"],
        ["📝 Доход", "💸 Расход"],
        ["📈 Бюджет", "📋 История"],
        ["📤 Экспорт в Excel", "❓ Инструкция пользователя"],
        ["🗑️ Сбросить всё"]
    ]
    if chat_id == ADMIN_ID:
        keyboard.append(["📊 Статистика"])
    return {"keyboard": keyboard, "resize_keyboard": True}

def two_column_keyboard(cats):
    keyboard = []
    for i in range(0, len(cats), 2):
        row = []
        row.append(cats[i])
        if i + 1 < len(cats):
            row.append(cats[i + 1])
        keyboard.append(row)
    keyboard.append(["🔙 Назад"])
    return {"keyboard": keyboard, "resize_keyboard": True}

def category_keyboard(trans_type):
    cats = list(CATEGORIES[trans_type].keys())
    return two_column_keyboard(cats)

def back_keyboard():
    return {"keyboard": [["🔙 Назад"]], "resize_keyboard": True}

def budget_keyboard_with_delete():
    return {"keyboard": [
        ["🗑️ Удалить лимит"],
        ["🔙 Назад"]
    ], "resize_keyboard": True}

def history_keyboard(page, total_pages):
    keyboard = []
    if page > 0:
        keyboard.append(["🔙 Назад"])
    if page < total_pages - 1:
        keyboard.append(["➡️ Дальше"])
    if page > 0 and page < total_pages - 1:
        keyboard = [["🔙 Назад", "➡️ Дальше"]]
    elif page > 0 and page == total_pages - 1:
        keyboard = [["🔙 Назад"]]
    elif page == 0 and page < total_pages - 1:
        keyboard = [["➡️ Дальше"]]
    keyboard.append(["🔙 Главное меню"])
    return {"keyboard": keyboard, "resize_keyboard": True}

def months_keyboard(months):
    keyboard = []
    for month in months:
        year, month_num = month.split("-")
        month_name = datetime(int(year), int(month_num), 1).strftime("%B %Y")
        keyboard.append([f"📆 {month_name}"])
    keyboard.append(["🔙 Назад"])
    return {"keyboard": keyboard, "resize_keyboard": True}

def confirm_delete_keyboard():
    return {"keyboard": [
        ["✅ Да, удалить всё"],
        ["❌ Нет, отмена"]
    ], "resize_keyboard": True}

def send_message(chat_id, text, keyboard=None):
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    data = {"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}
    if keyboard:
        data["reply_markup"] = json.dumps(keyboard)
    try:
        requests.post(url, data=data, timeout=5)
    except Exception as e:
        print(f"⚠️ Ошибка отправки: {e}")

# ===== ОБРАБОТЧИКИ =====
def handle_start(chat_id):
    create_user(chat_id)
    handle_help(chat_id)

def handle_help(chat_id):
    text1 = (
        "📖 *Инструкция пользователя (часть 1/3)*\n\n"
        "💰 *Финансовый помощник* помогает вести учёт доходов и расходов.\n\n"
        "🔹 *Как записать доход:*\n"
        "1. Нажмите «📝 Доход»\n"
        "2. Выберите категорию (например, «💰 Зарплата»)\n"
        "3. Введите сумму и описание\n"
        "   *Пример:* 15000 зарплата\n\n"
        "🔹 *Как записать расход:*\n"
        "1. Нажмите «💸 Расход»\n"
        "2. Выберите категорию (например, «🍔 Еда»)\n"
        "3. Введите сумму и описание\n"
        "   *Пример:* 500 обед\n\n"
        "📌 *Если меню не появилось — просто отправьте «Старт» ещё раз.*"
    )
    send_message(chat_id, text1, main_keyboard(chat_id))
    time.sleep(0.5)
    
    text2 = (
        "📖 *Инструкция пользователя (часть 2/3)*\n\n"
        "🔹 *Что такое бюджет?*\n"
        "Это лимит на категорию в месяц.\n"
        "Нажмите «📈 Бюджет», выберите категорию — бот покажет текущий лимит (если он есть).\n"
        "Вы можете установить новый лимит, изменить его или удалить.\n"
        "Если превысите лимит — бот предупредит.\n\n"
        "🔹 *Как посмотреть отчёт?*\n"
        "Нажмите «📊 Отчёт» — бот покажет текущий месяц.\n"
        "Нажмите «📅 Выбрать месяц» — можно посмотреть любой месяц."
    )
    send_message(chat_id, text2)
    time.sleep(0.5)
    
    text3 = (
        "📖 *Инструкция пользователя (часть 3/3)*\n\n"
        "🔹 *Как посмотреть историю?*\n"
        "Нажмите «📋 История» — бот покажет последние операции.\n"
        "Листайте с помощью кнопок «🔙 Назад» и «➡️ Дальше».\n\n"
        "🔹 *Как удалить все данные?*\n"
        "Нажмите «🗑️ Сбросить всё» — бот запросит подтверждение."
    )
    send_message(chat_id, text3, main_keyboard(chat_id))

def handle_stats(chat_id):
    if chat_id != ADMIN_ID:
        send_message(chat_id, "⛔ У вас нет доступа к этой команде.", main_keyboard(chat_id))
        return
    
    stats = get_stats()
    text = (
        "📊 *Статистика бота*\n\n"
        f"👥 Всего пользователей: {stats['total_users']}\n"
        f"📆 Активных за 7 дней: {stats['active_users']}\n"
        f"📝 Операций за месяц: {stats['total_ops']}\n"
        f"💰 Доходы (всех пользователей): {format_amount(stats['total_income'])} ₽\n"
        f"📉 Расходы (всех пользователей): {format_amount(stats['total_expense'])} ₽"
    )
    send_message(chat_id, text, main_keyboard(chat_id))

def handle_balance(chat_id):
    balance, income, expense = get_balance(chat_id)
    text = (
        f"💰 *Ваш баланс: {format_amount(balance)} ₽*\n\n"
        f"↗️ Доходы: {format_amount(income)} ₽\n"
        f"↘️ Расходы: {format_amount(expense)} ₽"
    )
    send_message(chat_id, text, main_keyboard(chat_id))

def handle_report_current(chat_id):
    data = get_report_for_month(chat_id, datetime.now().year, datetime.now().month)
    if not data:
        months = get_available_months(chat_id)
        if months:
            text = "📊 За этот месяц операций нет.\n\n📅 Выберите другой месяц:"
            send_message(chat_id, text, months_keyboard(months))
            user_states[chat_id] = {"action": "choose_month"}
        else:
            send_message(chat_id, "📊 За этот месяц операций нет.", main_keyboard(chat_id))
        return
    
    balance, income, expense = get_balance(chat_id)
    text = f"📊 *Отчёт за {datetime.now().strftime('%B %Y')}*\n\n"
    text += f"↗️ Доходы: {format_amount(income)} ₽\n"
    text += f"↘️ Расходы: {format_amount(expense)} ₽\n"
    text += f"💵 Свободно: {format_amount(balance)} ₽\n\n"
    
    expense_cats = {}
    for trans_type, category, amount in data:
        if trans_type == "expense":
            expense_cats[category] = expense_cats.get(category, 0) + amount
    
    if expense_cats:
        text += "🔹 *Расходы по категориям:*\n"
        for cat, amount in sorted(expense_cats.items(), key=lambda x: x[1], reverse=True):
            budget = get_budget(chat_id, cat)
            line = f"  • {cat}: {format_amount(amount)} ₽"
            if budget:
                line += f" (лимит: {format_amount(budget)} ₽)"
                if amount > budget:
                    line += f" ⚠️ превышен!"
            text += line + "\n"
    else:
        text += "🔹 Расходов нет\n"
    
    keyboard = main_keyboard(chat_id)
    keyboard["keyboard"].append(["📅 Выбрать месяц"])
    send_message(chat_id, text, keyboard)

def handle_report_by_month(chat_id, year, month):
    data = get_report_for_month(chat_id, year, month)
    income, expense = get_monthly_income_expense(chat_id, year, month)
    
    if not data:
        send_message(chat_id, f"📊 За {datetime(year, month, 1).strftime('%B %Y')} операций нет.", main_keyboard(chat_id))
        return
    
    balance = income - expense
    text = f"📊 *Отчёт за {datetime(year, month, 1).strftime('%B %Y')}*\n\n"
    text += f"↗️ Доходы: {format_amount(income)} ₽\n"
    text += f"↘️ Расходы: {format_amount(expense)} ₽\n"
    text += f"💵 Свободно: {format_amount(balance)} ₽\n\n"
    
    expense_cats = {}
    for trans_type, category, amount in data:
        if trans_type == "expense":
            expense_cats[category] = expense_cats.get(category, 0) + amount
    
    if expense_cats:
        text += "🔹 *Расходы по категориям:*\n"
        for cat, amount in sorted(expense_cats.items(), key=lambda x: x[1], reverse=True):
            budget = get_budget(chat_id, cat)
            line = f"  • {cat}: {format_amount(amount)} ₽"
            if budget:
                line += f" (лимит: {format_amount(budget)} ₽)"
                if amount > budget:
                    line += f" ⚠️ превышен!"
            text += line + "\n"
    else:
        text += "🔹 Расходов нет\n"
    
    send_message(chat_id, text, main_keyboard(chat_id))

def handle_history(chat_id, page=0):
    total = get_total_transactions(chat_id)
    if total == 0:
        send_message(chat_id, "📋 История пуста.", main_keyboard(chat_id))
        return
    
    per_page = 10
    total_pages = (total + per_page - 1) // per_page
    
    if page >= total_pages:
        page = total_pages - 1
    
    transactions = get_transactions_page(chat_id, page, per_page)
    
    text = f"📋 *История операций (стр. {page + 1}/{total_pages})*\n\n"
    
    for trans_type, category, amount, description, date in transactions:
        emoji = "↗️" if trans_type == "income" else "↘️"
        desc = f" ({description})" if description else ""
        text += f"{emoji} {category}: {format_amount(amount)} ₽{desc}\n"
    
    send_message(chat_id, text, history_keyboard(page, total_pages))
    
    user_states[chat_id] = {"action": "history", "page": page, "total_pages": total_pages}

def handle_budget(chat_id):
    cats = list(CATEGORIES["expense"].keys())
    send_message(chat_id, "📈 *Выберите категорию для бюджета:*", two_column_keyboard(cats))

def handle_choose_month(chat_id):
    months = get_available_months(chat_id)
    if not months:
        send_message(chat_id, "📅 Нет данных для выбора месяца.", main_keyboard(chat_id))
        return
    
    text = "📅 *Выберите месяц для отчёта:*"
    send_message(chat_id, text, months_keyboard(months))
    user_states[chat_id] = {"action": "choose_month"}

print("✅ Бот готов! Жду сообщения...")
print("=" * 50)

offset = 0
user_states = {}

# ===== ОСНОВНОЙ ЦИКЛ =====
while True:
    try:
        url = f"https://api.telegram.org/bot{TG_TOKEN}/getUpdates?offset={offset}&timeout=30"
        resp = requests.get(url)
        updates = resp.json().get("result", [])

        for update in updates:
            if "message" in update:
                msg = update["message"]
                chat_id = msg["chat"]["id"]
                text = msg.get("text", "").strip()

                if not get_user(chat_id):
                    create_user(chat_id)

                state = user_states.get(chat_id)

                # ============================================================
                # 1. ОБРАБОТКА КОМАНД
                # ============================================================
                if text == "/start":
                    user_states.pop(chat_id, None)
                    handle_start(chat_id)
                    offset = update["update_id"] + 1
                    continue

                if text == "/help" or text == "❓ Инструкция пользователя":
                    user_states.pop(chat_id, None)
                    handle_help(chat_id)
                    offset = update["update_id"] + 1
                    continue

                if text == "📊 Статистика":
                    handle_stats(chat_id)
                    offset = update["update_id"] + 1
                    continue

                if text == "📤 Экспорт в Excel":
                    user_states.pop(chat_id, None)
                    export_to_excel(chat_id)
                    offset = update["update_id"] + 1
                    continue

                if text == "🔙 Назад":
                    if state and state.get("action") == "history":
                        page = state.get("page", 0)
                        if page > 0:
                            handle_history(chat_id, page - 1)
                        else:
                            user_states.pop(chat_id, None)
                            send_message(chat_id, "🔙 Главное меню", main_keyboard(chat_id))
                    else:
                        user_states.pop(chat_id, None)
                        send_message(chat_id, "🔙 Главное меню", main_keyboard(chat_id))
                    offset = update["update_id"] + 1
                    continue

                if text == "➡️ Дальше":
                    if state and state.get("action") == "history":
                        page = state.get("page", 0) + 1
                        total_pages = state.get("total_pages", 1)
                        if page < total_pages:
                            handle_history(chat_id, page)
                        else:
                            send_message(chat_id, "📋 Это последняя страница.", main_keyboard(chat_id))
                    else:
                        send_message(chat_id, "Используйте кнопки меню 👇", main_keyboard(chat_id))
                    offset = update["update_id"] + 1
                    continue

                if text == "🔙 Главное меню":
                    user_states.pop(chat_id, None)
                    send_message(chat_id, "🔙 Главное меню", main_keyboard(chat_id))
                    offset = update["update_id"] + 1
                    continue

                if text == "💰 Баланс":
                    user_states.pop(chat_id, None)
                    handle_balance(chat_id)
                    offset = update["update_id"] + 1
                    continue

                if text == "📊 Отчёт":
                    user_states.pop(chat_id, None)
                    handle_report_current(chat_id)
                    offset = update["update_id"] + 1
                    continue

                if text == "📋 История":
                    user_states.pop(chat_id, None)
                    handle_history(chat_id, 0)
                    offset = update["update_id"] + 1
                    continue

                if text == "📝 Доход":
                    user_states[chat_id] = {"action": "income_select"}
                    send_message(chat_id, "📝 Выберите категорию доходов, либо нажмите 🔙 для выхода в главное меню", category_keyboard("income"))
                    offset = update["update_id"] + 1
                    continue

                if text == "💸 Расход":
                    user_states[chat_id] = {"action": "expense_select"}
                    send_message(chat_id, "💸 Выберите категорию расходов, либо нажмите 🔙 для выхода в главное меню", category_keyboard("expense"))
                    offset = update["update_id"] + 1
                    continue

                if text == "📈 Бюджет":
                    user_states[chat_id] = {"action": "budget_select"}
                    handle_budget(chat_id)
                    offset = update["update_id"] + 1
                    continue

                if text == "🗑️ Сбросить всё":
                    user_states[chat_id] = {"action": "confirm_delete"}
                    send_message(
                        chat_id,
                        "⚠️ *ВНИМАНИЕ!*\n\nВы действительно хотите удалить ВСЕ свои данные?\n"
                        "(доходы, расходы, бюджеты)\n\n"
                        "Это действие НЕЛЬЗЯ отменить!",
                        confirm_delete_keyboard()
                    )
                    offset = update["update_id"] + 1
                    continue

                if text == "✅ Да, удалить всё":
                    if state and state.get("action") == "confirm_delete":
                        delete_all_data(chat_id)
                        user_states.pop(chat_id, None)
                        send_message(chat_id, "🗑️ *Все данные успешно удалены!*\n\nМожете начинать вести учёт заново.", main_keyboard(chat_id))
                    else:
                        send_message(chat_id, "Используйте кнопки меню 👇", main_keyboard(chat_id))
                    offset = update["update_id"] + 1
                    continue

                if text == "❌ Нет, отмена":
                    if state and state.get("action") == "confirm_delete":
                        user_states.pop(chat_id, None)
                        send_message(chat_id, "✅ Отмена. Ваши данные сохранены.", main_keyboard(chat_id))
                    else:
                        send_message(chat_id, "Используйте кнопки меню 👇", main_keyboard(chat_id))
                    offset = update["update_id"] + 1
                    continue

                if text == "📅 Выбрать месяц":
                    user_states.pop(chat_id, None)
                    handle_choose_month(chat_id)
                    offset = update["update_id"] + 1
                    continue

                # ============================================================
                # 2. ОБРАБОТКА СОСТОЯНИЙ
                # ============================================================

                if state and state.get("action") == "income_select":
                    if text in CATEGORIES["income"]:
                        state["category"] = text
                        state["action"] = "income_amount"
                        example = CATEGORIES["income"][text]
                        send_message(chat_id, f"💰 Введите сумму и описание для *{text}*\nНапример: *15000 {example}*", back_keyboard())
                    else:
                        send_message(chat_id, "Выберите категорию из списка.", category_keyboard("income"))
                    offset = update["update_id"] + 1
                    continue

                if state and state.get("action") == "income_amount":
                    amount, description = extract_amount_and_desc(text)
                    if amount is not None:
                        add_transaction(chat_id, "income", state["category"], amount, description or "")
                        desc_text = f" ({description})" if description else ""
                        send_message(chat_id, f"✅ Доход записан!\n{state['category']}: {format_amount(amount)} ₽{desc_text}", main_keyboard(chat_id))
                        user_states.pop(chat_id, None)
                    else:
                        example = CATEGORIES["income"].get(state["category"], "сумма")
                        send_message(chat_id, f"Введите число и описание.\nНапример: *15000 {example}*", back_keyboard())
                    offset = update["update_id"] + 1
                    continue

                if state and state.get("action") == "expense_select":
                    if text in CATEGORIES["expense"]:
                        state["category"] = text
                        state["action"] = "expense_amount"
                        example = CATEGORIES["expense"][text]
                        send_message(chat_id, f"💸 Введите сумму и описание для *{text}*\nНапример: *500 {example}*", back_keyboard())
                    else:
                        send_message(chat_id, "Выберите категорию из списка.", category_keyboard("expense"))
                    offset = update["update_id"] + 1
                    continue

                if state and state.get("action") == "expense_amount":
                    amount, description = extract_amount_and_desc(text)
                    if amount is not None:
                        add_transaction(chat_id, "expense", state["category"], amount, description or "")

                        budget = get_budget(chat_id, state["category"])
                        if budget:
                            spent = get_monthly_expense(chat_id, state["category"])
                            if spent > budget:
                                send_message(
                                    chat_id,
                                    f"⚠️ *Превышен бюджет на «{state['category']}»!*\n"
                                    f"📉 Лимит: {format_amount(budget)} ₽\n"
                                    f"📈 Потрачено: {format_amount(spent)} ₽\n"
                                    f"🔥 Перерасход: {format_amount(spent - budget)} ₽",
                                    main_keyboard(chat_id)
                                )

                        desc_text = f" ({description})" if description else ""
                        send_message(
                            chat_id,
                            f"✅ Расход записан!\n{state['category']}: {format_amount(amount)} ₽{desc_text}",
                            main_keyboard(chat_id)
                        )
                        user_states.pop(chat_id, None)
                    else:
                        example = CATEGORIES["expense"].get(state["category"], "сумма")
                        send_message(
                            chat_id,
                            f"Введите число и описание.\nНапример: *500 {example}*",
                            back_keyboard()
                        )
                    offset = update["update_id"] + 1
                    continue

                if state and state.get("action") == "budget_select":
                    if text in CATEGORIES["expense"]:
                        state["category"] = text
                        current_limit = get_budget(chat_id, text)
                        if current_limit is not None:
                            msg = (
                                f"📈 *Бюджет для категории «{text}»*\n\n"
                                f"Текущий лимит: {format_amount(current_limit)} ₽\n\n"
                                f"Введите новый лимит (или нажмите «🗑️ Удалить лимит»):"
                            )
                            send_message(chat_id, msg, budget_keyboard_with_delete())
                            state["action"] = "budget_amount"
                        else:
                            msg = (
                                f"📈 *Бюджет для категории «{text}»*\n\n"
                                f"Лимит не установлен.\n\n"
                                f"Введите сумму лимита (или нажмите «🔙 Назад»):"
                            )
                            send_message(chat_id, msg, back_keyboard())
                            state["action"] = "budget_amount"
                    elif text == "🔙 Назад":
                        user_states.pop(chat_id, None)
                        send_message(chat_id, "🔙 Главное меню", main_keyboard(chat_id))
                    else:
                        send_message(chat_id, "Выберите категорию из списка.", two_column_keyboard(list(CATEGORIES["expense"].keys())))
                    offset = update["update_id"] + 1
                    continue

                if state and state.get("action") == "budget_amount":
                    if text == "🗑️ Удалить лимит":
                        delete_budget(chat_id, state["category"])
                        send_message(chat_id, f"🗑️ Лимит для категории *{state['category']}* удалён.", main_keyboard(chat_id))
                        user_states.pop(chat_id, None)
                    elif text == "🔙 Назад":
                        user_states.pop(chat_id, None)
                        send_message(chat_id, "🔙 Главное меню", main_keyboard(chat_id))
                    else:
                        try:
                            amount = float(text.replace(",", "."))
                            set_budget(chat_id, state["category"], amount)
                            send_message(chat_id, f"✅ Бюджет для *{state['category']}*: {format_amount(amount)} ₽", main_keyboard(chat_id))
                            user_states.pop(chat_id, None)
                        except ValueError:
                            send_message(chat_id, "❌ Введите число!", back_keyboard())
                    offset = update["update_id"] + 1
                    continue

                if state and state.get("action") == "choose_month":
                    if text.startswith("📆 "):
                        month_str = text.replace("📆 ", "")
                        try:
                            month_name, year_str = month_str.rsplit(" ", 1)
                            year = int(year_str)
                            month_num = datetime.strptime(month_name, "%B").month
                            handle_report_by_month(chat_id, year, month_num)
                            user_states.pop(chat_id, None)
                        except:
                            send_message(chat_id, "Ошибка при выборе месяца.", main_keyboard(chat_id))
                    else:
                        send_message(chat_id, "Выберите месяц из списка.", main_keyboard(chat_id))
                    offset = update["update_id"] + 1
                    continue

                # ============================================================
                # 3. НЕИЗВЕСТНАЯ КОМАНДА
                # ============================================================
                send_message(chat_id, "Используйте кнопки меню 👇", main_keyboard(chat_id))
                offset = update["update_id"] + 1

        time.sleep(2)

    except Exception as e:
        print(f"⚠️ Ошибка: {e}")
        time.sleep(5)
