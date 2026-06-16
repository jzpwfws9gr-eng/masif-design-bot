import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
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

HEADERS = ["المشارك", "الحارس", "اللاعب 1", "اللاعب 2", "اللاعب 3", "الكابتن"]


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


def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    row_fill = PatternFill("solid", fgColor="D9EAF7")
    white_font = Font(color="FFFFFF", bold=True, size=12)
    normal_font = Font(size=12)
    gray_font = Font(color="808080", size=12)
    thin = Side(style="thin", color="5B9BD5")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:F13"
    ws.sheet_view.rightToLeft = True

    widths = {
        "A": 18,
        "B": 18,
        "C": 18,
        "D": 18,
        "E": 18,
        "F": 18,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            cell.font = normal_font

            if cell.row == 1:
                cell.fill = header_fill
                cell.font = white_font
            elif cell.row % 2 == 0:
                cell.fill = row_fill

            if cell.value == "لم يشارك":
                cell.font = gray_font

    ws.row_dimensions[1].height = 28
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 24


def create_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "اليوم5"

    ws.append(HEADERS)

    for name in PARTICIPANTS:
        values = data.get(name, ["لم يشارك"] * 5)
        ws.append([name] + values)

    style_sheet(ws)
    wb.save(EXCEL_FILE)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "أرسل كذا:\n"
        "/اضافه_اليوم5\n\n"
        "فهد فارس|أوناي سيمون|داني أولمو|سالم الدوسري|داروين نونيز|داروين نونيز"
    )


async def add_day5(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lines = update.message.text.splitlines()[1:]
    data = {}

    for line in lines:
        line = line.strip()
        if not line or "|" not in line:
            continue

        parts = [p.strip() for p in line.split("|")]

        if len(parts) == 6:
            participant = parts[0]
            values = [normalize_name(x) for x in parts[1:]]
            data[participant] = values

    create_excel(data)

    with open(EXCEL_FILE, "rb") as file:
        await update.message.reply_document(
            document=file,
            filename=EXCEL_FILE,
            caption="تم إنشاء ملف اليوم الخامس منسق ✅"
        )


def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اضافه_اليوم5"), add_day5))

    app.run_polling()


if __name__ == "__main__":
    main()