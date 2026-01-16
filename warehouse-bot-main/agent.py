import os
import re
from openpyxl import load_workbook
from telegram import Update
from telegram.ext import (
    Application,
    MessageHandler,
    CommandHandler,
    ContextTypes,
    filters,
)

TOKEN = os.getenv("TOKEN")
EXCEL_FILE = "warehouse.xlsx"


# ---------- НОРМАЛИЗАЦИЯ ТЕКСТА ----------
def normalize(text):
    if text is None:
        return ""
    text = str(text).lower()
    text = text.replace("ё", "е")
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# ---------- ЗАГРУЗКА ИЗ EXCEL ----------
def load_items():
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    items = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue

        items.append({
            "name": str(row[0]),
            "quantity": row[1],
            "shelf": row[2],
            "cell": row[3],
            "passport": row[4],
            "category": row[5],
            "serial": row[6],
            "checked": row[7],
        })

    return items


ITEMS = load_items()


# ---------- ПОИСК ----------
def search_items(query):
    q = normalize(query)
    results = []

    for item in ITEMS:
        name_norm = normalize(item["name"])

        if q in name_norm or name_norm in q:
            results.append(item)

    return results


# ---------- /start ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет 👋\n"
        "Напиши номер детали или серийный номер — я проверю склад."
    )


# ---------- СООБЩЕНИЯ ----------
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    found = search_items(text)

    if not found:
        await update.message.reply_text("❌ Не найдено")
        return

    for item in found:
        msg = (
            f"✅ *{item['name']} есть в наличии*\n"
            f"📦 Полка: {item['shelf']}, ячейка: {item['cell']}\n"
            f"🔢 Количество: {item['quantity']}\n"
            f"📄 Паспорт: {'есть' if item['passport'] else 'нет'}\n"
            f"🆕 Категория: {item['category']}\n"
            f"🔑 Серийный номер: {item['serial']}\n"
            f"✔ Проверка: {'проверена' if item['checked'] else 'не проверена'}"
        )

        await update.message.reply_text(msg, parse_mode="Markdown")


# ---------- MAIN ----------
def main():
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    app.run_polling()


if __name__ == "__main__":
    main()
