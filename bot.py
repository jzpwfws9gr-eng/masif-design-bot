import os
from openpyxl import Workbook, load_workbook
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

TOKEN = os.getenv("BOT_TOKEN")
EXCEL_FILE = "fantasy.xlsx"

PARTICIPANTS = [
    "عادل عيد",
    "فهد فارس",
    "نواف فارس",
    "خالد عبدالرحمن",
    "محمد عبدالرحمن",
    "سلطان رباح",
    "فارس سالم",
    "عبدالرحمن سالم",
    "ممدوح غزاي",
    "محمد محسن",
    "طلال عبدالله",
    "مشعل غزاي",
]

HEADERS = ["الاسم", "الحارس", "اللاعب 1", "اللاعب 2", "اللاعب 3", "الكابتن"]


def create_or_open_excel():
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "اليوم5"
        ws.append(HEADERS)
        for name in PARTICIPANTS:
            ws.append([name, "", "", "", "", ""])
        wb.save(EXCEL_FILE)

    return wb


def normalize_name(name):
    fixes = {
        "سيمون": "أوناي سيمون",
        "اوناي سيمون": "أوناي سيمون",
        "اويارزابال": "أويارزابال",
        "اويارزبال": "أويارزابال",
        "اوزبال": "أويارزابال",
        "توريس": "فيران توريس",
        "نونيز": "داروين نونيز",
        "نوينز": "داروين نونيز",
        "مرموش": "عمر مرموش",
        "يامال": "لامين يامال",
        "رايا": "دافيد رايا",
        "دافيد": "دافيد رايا",
        "العويس": "محمد العويس",
        "اولمو": "داني أولمو",
        "أولمو": "داني أولمو",
    }
    name = name.strip()
    return fixes.get(name, name)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "أرسل التشكيلات كذا:\n"
        "/اضافه_اليوم5\n"
        "لم يشارك\tلم يشارك\tلم يشارك\tلم يشارك\tلم يشارك\n"
        "أوناي سيمون\tداني أولمو\tسالم الدوسري\tداروين نونيز\tداروين نونيز"
    )


async def add_day5(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    lines = text.splitlines()[1:]

    if not lines:
        await update.message.reply_text("ارسل البيانات تحت الأمر /اضافه_اليوم5")
        return

    wb = create_or_open_excel()
    ws = wb["اليوم5"]

    for i, participant in enumerate(PARTICIPANTS, start=2):
        if i - 2 < len(lines):
            parts = lines[i - 2].split("\t")

            if len(parts) == 5:
                values = [normalize_name(x) for x in parts]
            else:
                values = ["لم يشارك"] * 5
        else:
            values = ["لم يشارك"] * 5

        ws.cell(row=i, column=1).value = participant
        for col, value in enumerate(values, start=2):
            ws.cell(row=i, column=col).value = value

    wb.save(EXCEL_FILE)

    await update.message.reply_document(
        document=open(EXCEL_FILE, "rb"),
        filename=EXCEL_FILE,
        caption="تم تحديث اليوم الخامس ✅"
    )


def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))

    # لأن أوامر التليجرام العربية ما تشتغل بـ CommandHandler العادي
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اضافه_اليوم5"), add_day5))

    app.run_polling()


if __name__ == "__main__":
    main()