import os
from openpyxl import load_workbook
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, ContextTypes, filters

TOKEN = os.getenv(8551566060:AAFWo6JAdDoNqlkEq26CCxU1_OUO3oLE1Ac)
EXCEL_FILE = "warehouse.xlsx"

def normalize(text):
    if text is None:
        return ""
    return str(text).lower().replace(" ", "").replace("-", "")

def yes_no(value):
    v = normalize(value)
    return "да" if v in ("yes", "y", "true", "1", "да") else "нет"

def load_items():
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    items = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue

        items.append({
            "number": row[0],
            "quantity": row[1],
            "shelf": row[2],
            "location": row[3],
            "passport": row[4],
            "category": row[5],
            "serial": row[6],
            "check": row[8] if len(row) > 8 else None
        })

    return items

def find_items(query, items):
    q = normalize(query)
    results = []

    for item in items:
        if q in normalize(item["number"]) or q in normalize(item["serial"]):
            results.append(item)

    return results

def format_item(item):
    return (
        f"✅ {item['number']} есть в наличии\n"
        f"📦 Полка: {item['shelf']} | Ячейка: {item['location']}\n"
        f"🔢 Количество: {item['quantity']}\n"
        f"📄 Паспорт: {'есть' if yes_no(item['passport']) == 'да' else 'нет'}\n"
        f"🆕 Категория: {item['category']}\n"
        f"🔑 Серийный номер: {item['serial']}\n"
        f"✔️ Проверка: {yes_no(item['check'])}"
    )

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет 👋\n"
        "Напиши номер детали или серийный номер — я проверю склад."
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    items = load_items()
    found = find_items(text, items)

    if not found:
        await update.message.reply_text("❌ Не найдено")
        return

    reply = "\n\n".join(format_item(item) for item in found)
    await update.message.reply_text(reply)

def main():
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.run_polling()

if __name__ == "__main__":
    main()
