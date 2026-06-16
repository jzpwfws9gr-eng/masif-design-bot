import os
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

TOKEN = os.getenv("BOT_TOKEN")

PARTICIPANTS = [
    "عادل عيد", "فهد فارس", "نواف فارس", "خالد عبدالرحمن",
    "محمد عبدالرحمن", "سلطان رباح", "فارس سالم", "عبدالرحمن سالم",
    "ممدوح غزاي", "محمد محسن", "طلال عبدالله", "مشعل غزاي",
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
    green_fill = PatternFill("solid", fgColor="C6EFCE")

    white_font = Font(color="FFFFFF", bold=True, size=12)
    normal_font = Font(size=12)
    gray_font = Font(color="808080", size=12)
    bold_font = Font(bold=True, size=12)

    thin = Side(style="thin", color="5B9BD5")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=ws.max_column).column_letter}{ws.max_row}"
    ws.sheet_view.rightToLeft = True

    for col in range(1, ws.max_column + 1):
        letter = ws.cell(row=1, column=col).column_letter
        ws.column_dimensions[letter].width = 18

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            cell.font = normal_font

            if cell.row == 1:
                cell.fill = header_fill
                cell.font = white_font
            elif cell.row % 2 == 0:
                cell.fill = light_blue

            if cell.value == "لم يشارك":
                cell.font = gray_font

            if isinstance(cell.value, int) and cell.value > 0 and cell.row != 1:
                cell.fill = green_fill
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
        ws.append([name] + values + [0, 0, 0, 0, 0, 0])

    style_sheet(ws)
    wb.save(file_name)
    return file_name

def parse_results(text):
    goals, clean_sheets = [], []
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
        p1_points = 5 if p1 in goals else 0
        p2_points = 5 if p2 in goals else 0
        p3_points = 5 if p3 in goals else 0

        captain_points = 0
        if captain in goals:
            captain_points = 5
        if captain == keeper and keeper in clean_sheets:
            captain_points = 5

        total = keeper_points + p1_points + p2_points + p3_points + captain_points

        for col, val in zip(
            range(7, 13),
            [keeper_points, p1_points, p2_points, p3_points, captain_points, total]
        ):
            ws.cell(row=row, column=col).value = val

    style_sheet(ws)
    wb.save(file_name)
    return file_name

def create_overall_ranking():
    totals = {name: 0 for name in PARTICIPANTS}
    daily_wins = {name: 0 for name in PARTICIPANTS}
    days_found = []

    for day in range(1, 32):
        file_name = excel_file(day)
        if not os.path.exists(file_name):
            continue

        wb = load_workbook(file_name, data_only=True)
        sheet_name = f"اليوم{day}"
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        day_scores = {}

        for row in range(2, ws.max_row + 1):
            name = ws.cell(row=row, column=1).value
            score = ws.cell(row=row, column=12).value or 0

            if name in totals:
                totals[name] += int(score)
                day_scores[name] = int(score)

        if day_scores:
            max_score = max(day_scores.values())
            if max_score > 0:
                for name, score in day_scores.items():
                    if score == max_score:
                        daily_wins[name] += 1

        days_found.append(day)

    ranking = sorted(PARTICIPANTS, key=lambda n: totals[n], reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "الترتيب العام"
    ws.append(["المركز", "المشارك", "المجموع", "أسطورة اليوم"])

    for i, name in enumerate(ranking, start=1):
        ws.append([i, name, totals[name], daily_wins[name]])

    style_sheet(ws)
    file_name = "overall_ranking.xlsx"
    wb.save(file_name)

    return file_name, days_found

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "/اضافه 5\n"
        "/نتائج 5\n"
        "/الترتيب_العام"
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
            data[parts[0]] = [normalize_name(x) for x in parts[1:]]

    file_name = create_excel(day, data)

    with open(file_name, "rb") as file:
        await update.message.reply_document(
            document=file,
            filename=file_name,
            caption=f"تم إنشاء ملف اليوم {day} ✅"
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

async def overall(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_name, days_found = create_overall_ranking()

    if not days_found:
        await update.message.reply_text("ما لقيت أي ملفات أيام محسوبة.")
        return

    with open(file_name, "rb") as file:
        await update.message.reply_document(
            document=file,
            filename=file_name,
            caption=f"تم إنشاء الترتيب العام ✅\nالأيام المحسوبة: {', '.join(map(str, days_found))}"
        )

def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اضافه"), add_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج"), results_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الترتيب_العام"), overall))

    app.run_polling()

if __name__ == "__main__":
    main()