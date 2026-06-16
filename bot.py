import os
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

TOKEN = os.getenv("BOT_TOKEN")

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

HEADERS = [
    "المشارك", "الحارس", "اللاعب 1", "اللاعب 2", "اللاعب 3", "الكابتن",
    "نقاط الحارس", "نقاط لاعب 1", "نقاط لاعب 2", "نقاط لاعب 3",
    "نقاط الكابتن", "مجموع اليوم"
]


def excel_file(day):
    return f"fantasy_day_{day}.xlsx"


def get_day(text):
    match = re.search(r"(\d+)", text)
    return match.group(1) if match else "5"


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
    name = str(name).strip()
    return fixes.get(name, name)


def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    light_blue = PatternFill("solid", fgColor="D9EAF7")
    total_fill = PatternFill("solid", fgColor="DDEBF7")
    green_fill = PatternFill("solid", fgColor="C6EFCE")
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")

    white_font = Font(color="FFFFFF", bold=True, size=12)
    normal_font = Font(size=12)
    gray_font = Font(color="808080", size=12)
    bold_font = Font(bold=True, size=12)

    thin = Side(style="thin", color="5B9BD5")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{ws.max_row}"
    ws.sheet_view.rightToLeft = True

    widths = {
        "A": 18, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18,
        "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=12):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            cell.font = normal_font

            if cell.row == 1:
                cell.fill = header_fill
                cell.font = white_font
            elif cell.row % 2 == 0:
                cell.fill = light_blue

            if cell.column == 12 and cell.row != 1:
                cell.fill = total_fill
                cell.font = bold_font

            if cell.value == "لم يشارك":
                cell.font = gray_font

            if cell.row != 1 and cell.column in [7, 8, 9, 10] and isinstance(cell.value, int) and cell.value > 0:
                cell.fill = green_fill
                cell.font = bold_font

            if cell.row != 1 and cell.column == 11 and isinstance(cell.value, int) and cell.value > 0:
                cell.fill = yellow_fill
                cell.font = bold_font

    ws.row_dimensions[1].height = 30
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 24


def create_excel(day, data):
    file_name = excel_file(day)

    wb = Workbook()
    ws = wb.active
    ws.title = f"اليوم{day}"
    ws.append(HEADERS)

    for name in PARTICIPANTS:
        values = data.get(name, ["لم يشارك"] * 5)
        points = [0, 0, 0, 0, 0]
        total = 0
        ws.append([name] + values + points + [total])

    style_sheet(ws)
    wb.save(file_name)
    return file_name


def parse_results(text):
    goals = []
    clean_sheets = []
    mode = None

    for line in text.splitlines()[1:]:
        line = line.strip()
        if not line:
            continue

        if "الأهداف" in line or "اهداف" in line:
            mode = "goals"
            continue

        if "كلين" in line or "شيت" in line or "الكلين" in line:
            mode = "clean"
            continue

        if mode == "goals":
            goals.append(normalize_name(line))
        elif mode == "clean":
            clean_sheets.append(normalize_name(line))

    return goals, clean_sheets


def calculate_points(day, goals, clean_sheets):
    file_name = excel_file(day)

    if not os.path.exists(file_name):
        return None

    wb = load_workbook(file_name)
    ws = wb[f"اليوم{day}"]

    for row in range(2, ws.max_row + 1):
        keeper = normalize_name(ws.cell(row=row, column=2).value)
        p1 = normalize_name(ws.cell(row=row, column=3).value)
        p2 = normalize_name(ws.cell(row=row, column=4).value)
        p3 = normalize_name(ws.cell(row=row, column=5).value)
        captain = normalize_name(ws.cell(row=row, column=6).value)

        keeper_points = 5 if keeper in clean_sheets else 0
        p1_points = 10 if p1 in goals else 0
        p2_points = 10 if p2 in goals else 0
        p3_points = 10 if p3 in goals else 0

        captain_base = 0
        if captain in goals:
            captain_base = 10
        if captain == keeper and keeper in clean_sheets:
            captain_base = 5

        captain_points = captain_base * 2
        total = keeper_points + p1_points + p2_points + p3_points + captain_points

        ws.cell(row=row, column=7).value = keeper_points
        ws.cell(row=row, column=8).value = p1_points
        ws.cell(row=row, column=9).value = p2_points
        ws.cell(row=row, column=10).value = p3_points
        ws.cell(row=row, column=11).value = captain_points
        ws.cell(row=row, column=12).value = total

    style_sheet(ws)
    wb.save(file_name)
    return file_name


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "إضافة التشكيلات:\n"
        "/اضافه 5\n"
        "فهد فارس|أوناي سيمون|داني أولمو|سالم الدوسري|داروين نونيز|داروين نونيز\n\n"
        "إضافة النتائج:\n"
        "/نتائج 5\n\n"
        "الأهداف:\n"
        "داروين نونيز\n\n"
        "الكلين شيت:\n"
        "أوناي سيمون"
    )


async def add_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    day = get_day(text)
    lines = text.splitlines()[1:]
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

    file_name = create_excel(day, data)

    with open(file_name, "rb") as file:
        await update.message.reply_document(
            document=file,
            filename=file_name,
            caption=f"تم إنشاء ملف اليوم {day} مع أعمدة النقاط ✅"
        )


async def results_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    day = get_day(text)

    goals, clean_sheets = parse_results(text)
    file_name = calculate_points(day, goals, clean_sheets)

    if not file_name:
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}. أضف التشكيلات أولًا.")
        return

    with open(file_name, "rb") as file:
        await update.message.reply_document(
            document=file,
            filename=file_name,
            caption=f"تم حساب نقاط اليوم {day} ✅"
        )


def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اضافه"), add_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج"), results_day))

    app.run_polling()


if __name__ == "__main__":
    main()