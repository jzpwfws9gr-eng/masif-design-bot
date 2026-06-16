import os
import re
import json
import shutil
from datetime import datetime
from collections import Counter, defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, ContextTypes, filters

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

GOAL_POINTS = {1: 5, 2: 10, 3: 15, 4: 20, 5: 25, 6: 30}
LOCKED_FILE = "locked_days.json"
BACKUP_PREFIX = "backup_"

# -------------------- أدوات عامة --------------------

def excel_file(day):
    return f"fantasy_day_{day}.xlsx"


def get_numbers(text):
    return [int(x) for x in re.findall(r"\d+", text or "")]


def get_day(text, default=5):
    nums = get_numbers(text)
    return str(nums[0]) if nums else str(default)


def normalize_name(name):
    """بدون توحيد أسماء تلقائي: فقط تنظيف مسافات."""
    if name is None:
        return ""
    name = str(name).strip()
    name = re.sub(r"\s+", " ", name)
    return name


def is_no_participation(value):
    value = normalize_name(value)
    return value in ("", "لم يشارك")


def has_participated(row_values):
    return any(not is_no_participation(v) for v in row_values[:5])


def score_to_int(value):
    try:
        return int(value or 0)
    except Exception:
        return 0


def ordinal_day(day):
    names = {
        1: "الأول", 2: "الثاني", 3: "الثالث", 4: "الرابع", 5: "الخامس",
        6: "السادس", 7: "السابع", 8: "الثامن", 9: "التاسع", 10: "العاشر",
        11: "الحادي عشر", 12: "الثاني عشر", 13: "الثالث عشر", 14: "الرابع عشر", 15: "الخامس عشر",
    }
    try:
        return names.get(int(day), str(day))
    except Exception:
        return str(day)


def get_existing_days(start_day=1, end_day=31):
    days = []
    for filename in os.listdir("."):
        m = re.match(r"fantasy_day_(\d+)\.xlsx$", filename)
        if not m:
            continue
        day = int(m.group(1))
        if start_day <= day <= end_day:
            days.append(day)
    return sorted(days)


def load_locked_days():
    if not os.path.exists(LOCKED_FILE):
        return set()
    try:
        with open(LOCKED_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {str(x) for x in data.get("locked_days", [])}
    except Exception:
        return set()


def save_locked_days(days):
    with open(LOCKED_FILE, "w", encoding="utf-8") as f:
        json.dump({"locked_days": sorted([str(x) for x in days], key=lambda x: int(x))}, f, ensure_ascii=False, indent=2)


def is_locked(day):
    return str(day) in load_locked_days()


def current_data_files():
    files = []
    for filename in os.listdir("."):
        if (filename.startswith("fantasy_day_") and filename.endswith(".xlsx")) or filename == "overall_ranking.xlsx" or filename == "fantasy_dashboard.xlsx" or filename == LOCKED_FILE:
            files.append(filename)
    return sorted(files)


def backup_files(reason="auto", move=False, files=None):
    files = files if files is not None else current_data_files()
    files = [f for f in files if os.path.exists(f)]
    if not files:
        return None, []

    safe_reason = re.sub(r"[^A-Za-z0-9_\-\u0600-\u06FF]+", "_", reason).strip("_") or "auto"
    folder = f"{BACKUP_PREFIX}{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_reason}"
    os.makedirs(folder, exist_ok=True)

    for filename in files:
        target = os.path.join(folder, filename)
        if move:
            shutil.move(filename, target)
        else:
            shutil.copy2(filename, target)

    return folder, files


def latest_backup_folder():
    folders = [f for f in os.listdir(".") if os.path.isdir(f) and f.startswith(BACKUP_PREFIX)]
    if not folders:
        return None
    return sorted(folders)[-1]

# -------------------- تنسيق الإكسل --------------------

def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    light_blue = PatternFill("solid", fgColor="D9EAF7")
    green_fill = PatternFill("solid", fgColor="C6EFCE")
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")

    white_font = Font(color="FFFFFF", bold=True, size=12)
    normal_font = Font(size=12)
    gray_font = Font(color="808080", size=12)
    bold_font = Font(bold=True, size=12)

    thin = Side(style="thin", color="5B9BD5")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.sheet_view.rightToLeft = True

    if ws.max_row >= 1 and ws.max_column >= 1:
        last_col = ws.cell(row=1, column=ws.max_column).column_letter
        ws.auto_filter.ref = f"A1:{last_col}{ws.max_row}"

    for col in range(1, ws.max_column + 1):
        letter = ws.cell(row=1, column=col).column_letter
        ws.column_dimensions[letter].width = 20

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border
            cell.font = normal_font

            if cell.row == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = header_fill
                cell.font = white_font
            else:
                if isinstance(cell.value, str):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                if cell.row % 2 == 0:
                    cell.fill = light_blue

                if cell.value == "لم يشارك":
                    cell.font = gray_font

                if isinstance(cell.value, (int, float)) and cell.value > 0:
                    cell.fill = green_fill
                    cell.font = bold_font

    ws.row_dimensions[1].height = 30
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 24


def style_title_cell(cell):
    cell.font = Font(bold=True, size=16, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center", vertical="center")


def style_dashboard_sheet(ws):
    from openpyxl.utils import get_column_letter

    ws.sheet_view.rightToLeft = True

    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 22

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
# -------------------- ملفات الأيام --------------------

def create_blank_workbook(day):
    wb = Workbook()
    ws = wb.active
    ws.title = f"اليوم{day}"
    ws.append(HEADERS)
    for name in PARTICIPANTS:
        ws.append([name] + ["لم يشارك"] * 5 + [0, 0, 0, 0, 0, 0])
    style_sheet(ws)
    return wb


def get_or_create_day_workbook(day):
    file_name = excel_file(day)
    sheet_name = f"اليوم{day}"

    if os.path.exists(file_name):
        wb = load_workbook(file_name)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active
            ws.title = sheet_name
    else:
        wb = create_blank_workbook(day)
        ws = wb[sheet_name]

    # تأكد من وجود العناوين
    for idx, header in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=idx).value = header

    # تأكد من وجود كل المشاركين
    existing = {}
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=1).value
        if name:
            existing[normalize_name(name)] = row

    for name in PARTICIPANTS:
        if name not in existing:
            ws.append([name] + ["لم يشارك"] * 5 + [0, 0, 0, 0, 0, 0])

    return wb, ws, file_name


def participant_row(ws, participant):
    participant = normalize_name(participant)
    for row in range(2, ws.max_row + 1):
        if normalize_name(ws.cell(row=row, column=1).value) == participant:
            return row
    return None


def update_day_data(day, data):
    wb, ws, file_name = get_or_create_day_workbook(day)

    warnings = []
    for participant, values in data.items():
        row = participant_row(ws, participant)
        if not row:
            warnings.append(participant)
            continue

        for col, value in zip(range(2, 7), values):
            ws.cell(row=row, column=col).value = normalize_name(value)

        # عند تعديل تشكيلة مشارك، تصفر نقاطه حتى تعيد /نتائج
        for col in range(7, 13):
            ws.cell(row=row, column=col).value = 0

    style_sheet(ws)
    wb.save(file_name)
    return file_name, warnings


def read_day_rows(day, data_only=True):
    file_name = excel_file(day)
    sheet_name = f"اليوم{day}"
    if not os.path.exists(file_name):
        return []

    wb = load_workbook(file_name, data_only=data_only)
    if sheet_name not in wb.sheetnames:
        return []

    ws = wb[sheet_name]
    rows = []
    for row in range(2, ws.max_row + 1):
        item = {
            "participant": normalize_name(ws.cell(row=row, column=1).value),
            "keeper": normalize_name(ws.cell(row=row, column=2).value),
            "p1": normalize_name(ws.cell(row=row, column=3).value),
            "p2": normalize_name(ws.cell(row=row, column=4).value),
            "p3": normalize_name(ws.cell(row=row, column=5).value),
            "captain": normalize_name(ws.cell(row=row, column=6).value),
            "keeper_points": score_to_int(ws.cell(row=row, column=7).value),
            "p1_points": score_to_int(ws.cell(row=row, column=8).value),
            "p2_points": score_to_int(ws.cell(row=row, column=9).value),
            "p3_points": score_to_int(ws.cell(row=row, column=10).value),
            "captain_points": score_to_int(ws.cell(row=row, column=11).value),
            "total": score_to_int(ws.cell(row=row, column=12).value),
        }
        item["participated"] = has_participated([item["keeper"], item["p1"], item["p2"], item["p3"], item["captain"]])
        rows.append(item)
    return rows

# -------------------- النتائج والحساب --------------------

def parse_results(text):
    goals_count = defaultdict(int)
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
            if "|" in line:
                player, count = line.split("|", 1)
                player = normalize_name(player)
                try:
                    count = int(str(count).strip())
                except Exception:
                    count = 1
                goals_count[player] += count
            else:
                goals_count[normalize_name(line)] += 1

        elif mode == "clean":
            clean_sheets.append(normalize_name(line))

    clean_sheets = list(dict.fromkeys([x for x in clean_sheets if x]))
    goals_count = {k: v for k, v in goals_count.items() if k}
    return goals_count, clean_sheets


def goals_to_points(goals_count):
    return {player: GOAL_POINTS.get(count, count * 5) for player, count in goals_count.items()}


def lineup_names_for_day(day):
    names = set()
    for row in read_day_rows(day):
        if not row["participated"]:
            continue
        for key in ("keeper", "p1", "p2", "p3", "captain"):
            val = row[key]
            if val and val != "لم يشارك":
                names.add(val)
    return names


def validate_results_names(day, goals_count, clean_sheets):
    existing = lineup_names_for_day(day)
    goal_missing = [name for name in goals_count.keys() if name not in existing]
    clean_missing = [name for name in clean_sheets if name not in existing]
    return goal_missing, clean_missing


def calculate_points(day, goals_count, clean_sheets):
    file_name = excel_file(day)
    sheet_name = f"اليوم{day}"

    if not os.path.exists(file_name):
        return None

    wb = load_workbook(file_name)
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]

    goals_points = goals_to_points(goals_count)

    for row in range(2, ws.max_row + 1):
        keeper = normalize_name(ws.cell(row=row, column=2).value)
        p1 = normalize_name(ws.cell(row=row, column=3).value)
        p2 = normalize_name(ws.cell(row=row, column=4).value)
        p3 = normalize_name(ws.cell(row=row, column=5).value)
        captain = normalize_name(ws.cell(row=row, column=6).value)

        keeper_points = 5 if keeper in clean_sheets else 0
        p1_points = goals_points.get(p1, 0)
        p2_points = goals_points.get(p2, 0)
        p3_points = goals_points.get(p3, 0)

        captain_points = 0
        if captain == p1:
            captain_points = p1_points
        elif captain == p2:
            captain_points = p2_points
        elif captain == p3:
            captain_points = p3_points
        elif captain == keeper:
            captain_points = keeper_points

        total = keeper_points + p1_points + p2_points + p3_points + captain_points

        for col, val in zip(range(7, 13), [keeper_points, p1_points, p2_points, p3_points, captain_points, total]):
            ws.cell(row=row, column=col).value = val

    style_sheet(ws)
    wb.save(file_name)
    return file_name


def clear_day_points(day):
    file_name = excel_file(day)
    sheet_name = f"اليوم{day}"
    if not os.path.exists(file_name):
        return None

    wb = load_workbook(file_name)
    if sheet_name not in wb.sheetnames:
        return None

    ws = wb[sheet_name]
    for row in range(2, ws.max_row + 1):
        for col in range(7, 13):
            ws.cell(row=row, column=col).value = 0

    style_sheet(ws)
    wb.save(file_name)
    return file_name

# -------------------- تحليل البيانات --------------------

def collect_stats(start_day=1, end_day=31):
    days = get_existing_days(start_day, end_day)
    per_day = {}
    totals = {name: 0 for name in PARTICIPANTS}
    daily_wins = {name: 0 for name in PARTICIPANTS}
    participation_count = {name: 0 for name in PARTICIPANTS}
    zero_days = {name: 0 for name in PARTICIPANTS}
    scores_by_day = {name: {} for name in PARTICIPANTS}
    cumulative_by_day = {name: {} for name in PARTICIPANTS}

    captain_choice_count = Counter()
    captain_points_by_player = Counter()
    captain_points_by_participant = Counter()

    keeper_choice_count = Counter()
    keeper_points_by_keeper = Counter()
    keeper_points_by_participant = Counter()
    keeper_success_by_participant = Counter()

    player_choice_count = Counter()
    player_points = Counter()
    player_zero_selections = Counter()
    captain_selections_by_player = Counter()

    for day in days:
        rows = read_day_rows(day)
        day_scores = {}
        participants_today = 0
        zero_today = 0

        for row in rows:
            name = row["participant"]
            if name not in totals:
                continue

            total = row["total"]
            totals[name] += total
            scores_by_day[name][day] = total
            day_scores[name] = total

            if row["participated"]:
                participants_today += 1
                participation_count[name] += 1
                if total == 0:
                    zero_days[name] += 1
                    zero_today += 1

                keeper = row["keeper"]
                if keeper and keeper != "لم يشارك":
                    keeper_choice_count[keeper] += 1
                    keeper_points_by_keeper[keeper] += row["keeper_points"]
                    keeper_points_by_participant[name] += row["keeper_points"]
                    if row["keeper_points"] > 0:
                        keeper_success_by_participant[name] += 1

                captain = row["captain"]
                if captain and captain != "لم يشارك":
                    captain_choice_count[captain] += 1
                    captain_selections_by_player[captain] += 1
                    captain_points_by_player[captain] += row["captain_points"]
                    captain_points_by_participant[name] += row["captain_points"]

                for p_key, pts_key in (("p1", "p1_points"), ("p2", "p2_points"), ("p3", "p3_points")):
                    player = row[p_key]
                    pts = row[pts_key]
                    if player and player != "لم يشارك":
                        player_choice_count[player] += 1
                        player_points[player] += pts
                        if pts == 0:
                            player_zero_selections[player] += 1

        if day_scores:
            max_score = max(day_scores.values())
            winners = [n for n, s in day_scores.items() if s == max_score and max_score > 0]
            for winner in winners:
                daily_wins[winner] += 1

            scores_values = list(day_scores.values())
            per_day[day] = {
                "scores": day_scores,
                "max_score": max_score,
                "winners": winners,
                "participants": participants_today,
                "zeros": zero_today,
                "avg": round(sum(scores_values) / len(scores_values), 2) if scores_values else 0,
                "min_score": min(scores_values) if scores_values else 0,
            }

        for name in PARTICIPANTS:
            cumulative_by_day[name][day] = totals[name]

    ranking = sorted(PARTICIPANTS, key=lambda n: (totals[n], daily_wins[n]), reverse=True)

    return {
        "days": days,
        "per_day": per_day,
        "totals": totals,
        "daily_wins": daily_wins,
        "participation_count": participation_count,
        "zero_days": zero_days,
        "scores_by_day": scores_by_day,
        "cumulative_by_day": cumulative_by_day,
        "ranking": ranking,
        "captain_choice_count": captain_choice_count,
        "captain_points_by_player": captain_points_by_player,
        "captain_points_by_participant": captain_points_by_participant,
        "keeper_choice_count": keeper_choice_count,
        "keeper_points_by_keeper": keeper_points_by_keeper,
        "keeper_points_by_participant": keeper_points_by_participant,
        "keeper_success_by_participant": keeper_success_by_participant,
        "player_choice_count": player_choice_count,
        "player_points": player_points,
        "player_zero_selections": player_zero_selections,
        "captain_selections_by_player": captain_selections_by_player,
    }


def create_overall_ranking(start_day=1, end_day=31):
    stats = collect_stats(start_day, end_day)
    days_found = stats["days"]

    wb = Workbook()
    ws = wb.active
    ws.title = "الترتيب العام"
    ws.append(["المركز", "المشارك", "المجموع", "أسطورة اليوم"])

    current_rank = 0
    last_score = None
    real_index = 0

    for name in stats["ranking"]:
        real_index += 1
        score = stats["totals"][name]
        if score != last_score:
            current_rank = real_index
            last_score = score
        ws.append([current_rank, name, score, stats["daily_wins"][name]])

    style_sheet(ws)
    file_name = "overall_ranking.xlsx"
    wb.save(file_name)
    return file_name, days_found, stats


def build_ranking_text(stats, start_day=1, end_day=31):
    days = stats["days"]
    if not days:
        return "ما فيه أيام محسوبة."

    medals = ["🥇", "🥈", "🥉"]
    title = f"🏆 الترتيب العام لفانتزي المصيف 2026 بعد اليوم {max(days)} 🏆"
    lines = [title, ""]

    for idx, name in enumerate(stats["ranking"], start=1):
        score = stats["totals"][name]
        legends = stats["daily_wins"][name]
        icon = medals[idx - 1] if idx <= 3 else f"{idx}."
        if legends:
            lines.append(f"{icon} {name} — {score} نقطة | أسطورة اليوم: {legends}")
        else:
            lines.append(f"{icon} {name} — {score} نقطة")

    return "\n".join(lines)

# -------------------- الداشبورد --------------------

def add_bar_chart(ws, title, data_min_col, data_max_col, data_min_row, data_max_row, cats_col, cats_min_row, cats_max_row, anchor):
    if data_max_row < data_min_row:
        return
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = "الأسماء"
    chart.x_axis.title = "النقاط"
    data = Reference(ws, min_col=data_min_col, max_col=data_max_col, min_row=data_min_row, max_row=data_max_row)
    cats = Reference(ws, min_col=cats_col, min_row=cats_min_row, max_row=cats_max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 14
    ws.add_chart(chart, anchor)


def add_column_chart(ws, title, data_col, header_row, min_row, max_row, cats_col, anchor):
    if max_row < min_row:
        return
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    data = Reference(ws, min_col=data_col, min_row=header_row, max_row=max_row)
    cats = Reference(ws, min_col=cats_col, min_row=min_row, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 14
    ws.add_chart(chart, anchor)


def add_line_chart(ws, title, min_col, max_col, header_row, min_row, max_row, cats_col, anchor):
    if max_row < min_row or max_col < min_col:
        return
    chart = LineChart()
    chart.title = title
    chart.y_axis.title = "المجموع"
    chart.x_axis.title = "اليوم"
    data = Reference(ws, min_col=min_col, max_col=max_col, min_row=header_row, max_row=max_row)
    cats = Reference(ws, min_col=cats_col, min_row=min_row, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 10
    chart.width = 20
    ws.add_chart(chart, anchor)


def create_dashboard(start_day=1, end_day=31):
    from collections import defaultdict, Counter
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.chart import BarChart, LineChart, Reference

    stats = collect_stats(start_day, end_day)
    days = stats["days"]
    file_name = "fantasy_dashboard.xlsx"

    # ألوان ستايل فاتح حديث قريب من الداشبوردات
    BG = "F4F7FB"
    NAVY = "253858"
    BLUE = "4F7DF3"
    PURPLE = "7C3AED"
    GOLD = "F2B705"
    GREEN = "12B886"
    RED = "FA5252"
    TEXT = "1F2937"
    MUTED = "6B7280"
    WHITE = "FFFFFF"
    LIGHT_BLUE = "EAF1FF"
    LIGHT_PURPLE = "F1ECFF"
    LIGHT_GOLD = "FFF6D6"
    GRID = "D9E2EF"

    wb = Workbook()
    wb.remove(wb.active)

    def safe_pct(part, total):
        return f"{round((part / total) * 100, 1)}%" if total else "0%"

    def pct_number(part, total):
        return round((part / total) * 100, 1) if total else 0

    def top_key(counter, default="-"):
        if not counter:
            return default
        key, value = counter.most_common(1)[0]
        return key if value else default

    def top_value(counter):
        if not counter:
            return 0
        return counter.most_common(1)[0][1]

    def sheet_base(ws, title=None):
        ws.sheet_view.rightToLeft = True
        ws.freeze_panes = "A2"
        ws.sheet_properties.tabColor = BLUE

        for i in range(1, 20):
            ws.column_dimensions[get_column_letter(i)].width = 18

        for row in range(1, 80):
            ws.row_dimensions[row].height = 22

        if title:
            ws["A1"] = title
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
            c = ws["A1"]
            c.font = Font(bold=True, size=18, color=WHITE)
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 34

    def style_range_table(ws, header_row=1, min_row=1, max_row=None, min_col=1, max_col=None):
        max_row = max_row or ws.max_row
        max_col = max_col or ws.max_column
        thin = Side(style="thin", color=GRID)
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(cell.value, str):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.font = Font(size=11, color=TEXT)

                if cell.row == header_row:
                    cell.fill = PatternFill("solid", fgColor=NAVY)
                    cell.font = Font(bold=True, size=11, color=WHITE)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif cell.row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F8FAFC")
                else:
                    cell.fill = PatternFill("solid", fgColor=WHITE)

        ws.row_dimensions[header_row].height = 28

    def set_card(ws, cell_range, title, value, color=BLUE):
        ws.merge_cells(cell_range)
        c = ws[cell_range.split(":")[0]]
        c.value = f"{title}\n{value}"
        c.font = Font(bold=True, size=14, color=WHITE)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin = Side(style="thin", color=WHITE)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def add_section_title(ws, row, col, title, width=3, color=NAVY):
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + width - 1)
        c = ws.cell(row=row, column=col)
        c.value = title
        c.font = Font(bold=True, size=13, color=WHITE)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 28

    def add_bar(ws, title, data_col, header_row, min_row, max_row, cats_col, anchor, width=13, height=7):
        if max_row < min_row:
            return
        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = title
        chart.y_axis.title = ""
        chart.x_axis.title = ""
        data = Reference(ws, min_col=data_col, min_row=header_row, max_row=max_row)
        cats = Reference(ws, min_col=cats_col, min_row=min_row, max_row=max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = width
        chart.height = height
        ws.add_chart(chart, anchor)

    def add_col(ws, title, data_col, header_row, min_row, max_row, cats_col, anchor, width=13, height=7):
        if max_row < min_row:
            return
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = title
        chart.y_axis.title = ""
        chart.x_axis.title = ""
        data = Reference(ws, min_col=data_col, min_row=header_row, max_row=max_row)
        cats = Reference(ws, min_col=cats_col, min_row=min_row, max_row=max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = width
        chart.height = height
        ws.add_chart(chart, anchor)

    def add_line(ws, title, min_col, max_col, header_row, min_row, max_row, cats_col, anchor, width=18, height=8):
        if max_row < min_row or max_col < min_col:
            return
        chart = LineChart()
        chart.title = title
        chart.y_axis.title = "النقاط"
        chart.x_axis.title = "اليوم"
        data = Reference(ws, min_col=min_col, max_col=max_col, min_row=header_row, max_row=max_row)
        cats = Reference(ws, min_col=cats_col, min_row=min_row, max_row=max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = width
        chart.height = height
        ws.add_chart(chart, anchor)

    # حسابات إضافية للتوضيح
    active_scores_by_participant = {name: {} for name in PARTICIPANTS}
    player_participants = defaultdict(set)
    player_days_by_participant = defaultdict(lambda: defaultdict(list))
    player_captain_by_participant = defaultdict(Counter)
    captain_participants = defaultdict(set)
    keeper_participants = defaultdict(set)

    day_summary = {}
    for day in days:
        rows = read_day_rows(day)
        active_scores = []
        total_day_points = 0

        for row in rows:
            name = row["participant"]
            total = row["total"]
            total_day_points += total

            if row["participated"]:
                active_scores.append(total)
                active_scores_by_participant[name][day] = total

                keeper = row["keeper"]
                if keeper and keeper != "لم يشارك":
                    keeper_participants[keeper].add(name)

                captain = row["captain"]
                if captain and captain != "لم يشارك":
                    captain_participants[captain].add(name)

                for p_key in ("p1", "p2", "p3"):
                    player = row[p_key]
                    if player and player != "لم يشارك":
                        player_participants[player].add(name)
                        player_days_by_participant[player][name].append(day)
                        if captain == player:
                            player_captain_by_participant[player][name] += 1

        info = stats["per_day"].get(day, {})
        day_summary[day] = {
            "participants": info.get("participants", 0),
            "max_score": info.get("max_score", 0),
            "winners": info.get("winners", []),
            "avg_active": round(sum(active_scores) / len(active_scores), 2) if active_scores else 0,
            "total_points": total_day_points,
            "min_active": min(active_scores) if active_scores else 0,
        }

    ranking = stats["ranking"]
    leader = ranking[0] if ranking else "-"
    runner = ranking[1] if len(ranking) > 1 else None
    leader_points = stats["totals"].get(leader, 0)
    gap = leader_points - stats["totals"].get(runner, 0) if runner else 0

    most_legend = "-"
    if PARTICIPANTS:
        most_legend = max(PARTICIPANTS, key=lambda n: stats["daily_wins"][n])
        if stats["daily_wins"][most_legend] == 0:
            most_legend = "-"

    highest_daily = 0
    highest_daily_name = "-"
    highest_daily_day = "-"
    for day, info in day_summary.items():
        if info["max_score"] > highest_daily:
            highest_daily = info["max_score"]
            highest_daily_name = "، ".join(info["winners"]) or "-"
            highest_daily_day = day

    best_player = max(stats["player_points"], key=lambda p: stats["player_points"][p]) if stats["player_points"] else "-"
    best_captain = max(stats["captain_points_by_player"], key=lambda p: stats["captain_points_by_player"][p]) if stats["captain_points_by_player"] else "-"
    best_keeper = max(stats["keeper_points_by_keeper"], key=lambda p: stats["keeper_points_by_keeper"][p]) if stats["keeper_points_by_keeper"] else "-"
    most_selected_player = top_key(stats["player_choice_count"])
    most_selected_captain = top_key(stats["captain_choice_count"])

    # 1) لوحة عامة
    ws = wb.create_sheet("لوحة عامة")
    sheet_base(ws)
    ws.sheet_properties.tabColor = PURPLE

    ws["A1"] = "لوحة فانتزي المصيف 2026"
    ws.merge_cells("A1:L1")
    ws["A1"].font = Font(bold=True, size=22, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=PURPLE)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 42

    set_card(ws, "A3:C5", "🏆 المتصدر", f"{leader}\n{leader_points} نقطة", PURPLE)
    set_card(ws, "D3:F5", "👥 المشاركون", len(PARTICIPANTS), BLUE)
    set_card(ws, "G3:I5", "📅 الأيام", len(days), GREEN)
    set_card(ws, "J3:L5", "🔥 أعلى نتيجة يومية", f"{highest_daily_name}\n{highest_daily} - اليوم {highest_daily_day}", GOLD)

    set_card(ws, "A7:C9", "⭐ أكثر أسطورة يوم", f"{most_legend}\n{stats['daily_wins'].get(most_legend, 0)} مرات", NAVY)
    set_card(ws, "D7:F9", "⚽ أفضل لاعب نقاط", f"{best_player}\n{stats['player_points'].get(best_player, 0)}", GREEN)
    set_card(ws, "G7:I9", "👑 أفضل كابتن", f"{best_captain}\n{stats['captain_points_by_player'].get(best_captain, 0)}", PURPLE)
    set_card(ws, "J7:L9", "🧤 أفضل حارس", f"{best_keeper}\n{stats['keeper_points_by_keeper'].get(best_keeper, 0)}", BLUE)

    add_section_title(ws, 11, 1, "أفضل 5 في الترتيب", 3, NAVY)
    ws.cell(row=12, column=1).value = "المشارك"
    ws.cell(row=12, column=2).value = "النقاط"
    ws.cell(row=12, column=3).value = "الفارق عن الأول"
    for i, name in enumerate(ranking[:5], start=13):
        ws.cell(row=i, column=1).value = name
        ws.cell(row=i, column=2).value = stats["totals"][name]
        ws.cell(row=i, column=3).value = leader_points - stats["totals"][name]
    style_range_table(ws, header_row=12, min_row=12, max_row=17, min_col=1, max_col=3)

    add_section_title(ws, 11, 5, "أفضل الاختيارات", 3, NAVY)
    pick_rows = [
        ["أكثر لاعب اختيارًا", most_selected_player, top_value(stats["player_choice_count"])],
        ["أكثر كابتن اختيارًا", most_selected_captain, top_value(stats["captain_choice_count"])],
        ["أفضل لاعب نقاط", best_player, stats["player_points"].get(best_player, 0)],
        ["أفضل كابتن نقاط", best_captain, stats["captain_points_by_player"].get(best_captain, 0)],
        ["أفضل حارس نقاط", best_keeper, stats["keeper_points_by_keeper"].get(best_keeper, 0)],
    ]
    ws.cell(row=12, column=5).value = "المؤشر"
    ws.cell(row=12, column=6).value = "الاسم"
    ws.cell(row=12, column=7).value = "القيمة"
    for r, row in enumerate(pick_rows, start=13):
        for c, value in enumerate(row, start=5):
            ws.cell(row=r, column=c).value = value
    style_range_table(ws, header_row=12, min_row=12, max_row=17, min_col=5, max_col=7)

    # بيانات الرسوم داخل اللوحة
    chart_row = 21
    ws.cell(row=chart_row, column=1).value = "المشارك"
    ws.cell(row=chart_row, column=2).value = "النقاط"
    for idx, name in enumerate(ranking[:8], start=chart_row + 1):
        ws.cell(row=idx, column=1).value = name
        ws.cell(row=idx, column=2).value = stats["totals"][name]
    add_bar(ws, "ترتيب أفضل المشاركين", 2, chart_row, chart_row + 1, chart_row + min(8, len(ranking)), 1, "A29", width=15, height=8)

    day_chart_row = 21
    ws.cell(row=day_chart_row, column=5).value = "اليوم"
    ws.cell(row=day_chart_row, column=6).value = "مجموع النقاط"
    ws.cell(row=day_chart_row, column=7).value = "متوسط المشاركين"
    for idx, day in enumerate(days, start=day_chart_row + 1):
        ws.cell(row=idx, column=5).value = day
        ws.cell(row=idx, column=6).value = day_summary[day]["total_points"]
        ws.cell(row=idx, column=7).value = day_summary[day]["avg_active"]
    add_col(ws, "مجموع نقاط كل يوم", 6, day_chart_row, day_chart_row + 1, day_chart_row + len(days), 5, "E29", width=15, height=8)

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 2) الترتيب العام
    ws = wb.create_sheet("الترتيب العام")
    sheet_base(ws, "الترتيب العام")
    ws.append(["المركز", "المشارك", "النقاط", "الفارق عن المتصدر", "أفضل يوم", "نقاط أفضل يوم", "مرات أسطورة اليوم", "عدد المشاركات"])
    current_rank = 0
    last_score = None
    real_index = 0
    for name in ranking:
        real_index += 1
        score = stats["totals"][name]
        if score != last_score:
            current_rank = real_index
            last_score = score

        active = active_scores_by_participant[name]
        if active:
            best_day = max(active, key=lambda d: active[d])
            best_score = active[best_day]
        else:
            best_day, best_score = "-", 0

        ws.append([
            current_rank,
            name,
            score,
            leader_points - score,
            best_day,
            best_score,
            stats["daily_wins"][name],
            stats["participation_count"][name],
        ])
    style_range_table(ws, header_row=2, min_row=2)
    if ws.max_row >= 3:
        add_bar(ws, "النقاط حسب الترتيب", 3, 2, 3, ws.max_row, 2, "J3", width=15, height=8)

    # 3) تطور النقاط
    ws = wb.create_sheet("تطور النقاط")
    sheet_base(ws, "تطور النقاط")
    ws.append(["اليوم"] + ranking)
    for day in days:
        ws.append([day] + [stats["cumulative_by_day"][name].get(day, 0) for name in ranking])
    style_range_table(ws, header_row=2, min_row=2)
    if days and ranking:
        add_line(ws, "تطور مجموع النقاط يومًا بعد يوم", 2, 1 + len(ranking), 2, 3, ws.max_row, 1, "A16", width=22, height=10)

    # 4) تحليل الأيام
    ws = wb.create_sheet("تحليل الأيام")
    sheet_base(ws, "تحليل الأيام")
    ws.append(["اليوم", "عدد المشاركين", "أعلى نقاط", "أسطورة اليوم", "متوسط نقاط المشاركين", "مجموع نقاط اليوم"])
    for day in days:
        info = day_summary[day]
        ws.append([
            day,
            info["participants"],
            info["max_score"],
            "، ".join(info["winners"]) or "-",
            info["avg_active"],
            info["total_points"],
        ])
    style_range_table(ws, header_row=2, min_row=2)
    if ws.max_row >= 3:
        add_col(ws, "أعلى نقاط كل يوم", 3, 2, 3, ws.max_row, 1, "H3")
        add_col(ws, "متوسط نقاط المشاركين", 5, 2, 3, ws.max_row, 1, "H20")

    # 5) تحليل المشاركين
    ws = wb.create_sheet("تحليل المشاركين")
    sheet_base(ws, "تحليل المشاركين")
    ws.append(["المشارك", "المجموع", "أفضل يوم", "نقاط أفضل يوم", "أسوأ يوم", "نقاط أسوأ يوم", "متوسط نقاطه", "مرات أسطورة اليوم", "نسبة المشاركة"])
    total_days = len(days)
    for name in ranking:
        active = active_scores_by_participant[name]
        if active:
            best_day = max(active, key=lambda d: active[d])
            worst_day = min(active, key=lambda d: active[d])
            best_score = active[best_day]
            worst_score = active[worst_day]
            avg = round(sum(active.values()) / len(active), 2)
        else:
            best_day = worst_day = "-"
            best_score = worst_score = 0
            avg = 0

        pc = stats["participation_count"][name]
        ws.append([
            name,
            stats["totals"][name],
            best_day,
            best_score,
            worst_day,
            worst_score,
            avg,
            stats["daily_wins"][name],
            safe_pct(pc, total_days),
        ])
    style_range_table(ws, header_row=2, min_row=2)
    if ws.max_row >= 3:
        add_bar(ws, "مجموع نقاط المشاركين", 2, 2, 3, ws.max_row, 1, "K3", width=15, height=8)

    # 6) تحليل الكباتن
    ws = wb.create_sheet("تحليل الكباتن")
    sheet_base(ws, "تحليل الكباتن")
    ws.append(["اللاعب", "مرات كابتن", "عدد المشاركين الذين اختاروه كابتن", "نقاط كابتن", "متوسط نقاط كابتن"])
    captains = sorted(stats["captain_choice_count"].keys(), key=lambda p: (stats["captain_points_by_player"][p], stats["captain_choice_count"][p]), reverse=True)
    for player in captains:
        count = stats["captain_choice_count"][player]
        points = stats["captain_points_by_player"][player]
        ws.append([
            player,
            count,
            len(captain_participants[player]),
            points,
            round(points / count, 2) if count else 0,
        ])
    style_range_table(ws, header_row=2, min_row=2)
    if ws.max_row >= 3:
        add_col(ws, "أكثر لاعب تم اختياره كابتن", 2, 2, 3, min(ws.max_row, 17), 1, "G3")
        add_col(ws, "أفضل كابتن بالنقاط", 4, 2, 3, min(ws.max_row, 17), 1, "G20")

    # 7) تحليل الحراس
    ws = wb.create_sheet("تحليل الحراس")
    sheet_base(ws, "تحليل الحراس")
    ws.append(["الحارس", "إجمالي الاختيارات", "عدد المشاركين الذين اختاروه", "مرات الكلين شيت", "نقاط الحارس", "نسبة نجاح الحارس"])
    keepers = sorted(stats["keeper_choice_count"].keys(), key=lambda k: (stats["keeper_points_by_keeper"][k], stats["keeper_choice_count"][k]), reverse=True)
    for keeper in keepers:
        count = stats["keeper_choice_count"][keeper]
        points = stats["keeper_points_by_keeper"][keeper]
        clean_count = int(points / 5) if points else 0
        ws.append([
            keeper,
            count,
            len(keeper_participants[keeper]),
            clean_count,
            points,
            safe_pct(clean_count, count),
        ])
    style_range_table(ws, header_row=2, min_row=2)
    if ws.max_row >= 3:
        add_col(ws, "أكثر حارس تم اختياره", 2, 2, 3, min(ws.max_row, 17), 1, "H3")
        add_col(ws, "أفضل الحراس بالنقاط", 5, 2, 3, min(ws.max_row, 17), 1, "H20")

    # 8) تحليل اللاعبين
    ws = wb.create_sheet("تحليل اللاعبين")
    sheet_base(ws, "تحليل اللاعبين")
    ws.append(["اللاعب", "إجمالي الاختيارات", "عدد المشاركين الذين اختاروه", "مرات كابتن", "نقاط كلاعب", "نقاط كابتن", "إجمالي نقاطه"])
    players = sorted(stats["player_choice_count"].keys(), key=lambda p: (stats["player_points"][p] + stats["captain_points_by_player"][p], stats["player_choice_count"][p]), reverse=True)
    for player in players:
        player_points = stats["player_points"][player]
        captain_points = stats["captain_points_by_player"][player]
        ws.append([
            player,
            stats["player_choice_count"][player],
            len(player_participants[player]),
            stats["captain_selections_by_player"][player],
            player_points,
            captain_points,
            player_points + captain_points,
        ])
    style_range_table(ws, header_row=2, min_row=2)
    if ws.max_row >= 3:
        add_col(ws, "أكثر اللاعبين اختيارًا", 2, 2, 3, min(ws.max_row, 17), 1, "I3")
        add_col(ws, "أكثر اللاعبين نقاطًا", 7, 2, 3, min(ws.max_row, 17), 1, "I20")

    # 9) تفصيل اختيارات اللاعبين
    ws = wb.create_sheet("تفصيل اختيارات اللاعبين")
    sheet_base(ws, "تفصيل اختيارات اللاعبين")
    ws.append(["اللاعب", "المشارك", "الأيام المختار فيها", "عدد الاختيارات", "مرات كابتن"])
    detail_rows = []
    for player in sorted(player_days_by_participant.keys()):
        for participant in sorted(player_days_by_participant[player].keys(), key=lambda n: PARTICIPANTS.index(n) if n in PARTICIPANTS else 999):
            player_days = sorted(player_days_by_participant[player][participant])
            detail_rows.append([
                player,
                participant,
                "، ".join([str(d) for d in player_days]),
                len(player_days),
                player_captain_by_participant[player][participant],
            ])
    detail_rows.sort(key=lambda r: (r[0], r[1]))
    for row in detail_rows:
        ws.append(row)
    style_range_table(ws, header_row=2, min_row=2)

    # ضبط عرض الأعمدة لكل الصفحات
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
        for col in range(1, ws.max_column + 1):
            letter = get_column_letter(col)
            ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 12, 18)
        for row in range(1, ws.max_row + 1):
            ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 18, 22)

    wb.save(file_name)
    return file_name, stats

# -------------------- أوامر تيليجرام --------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "الأوامر الأساسية:\n"
        "/اضافه 5\n"
        "/نتائج 5\n"
        "/الترتيب_العام\n"
        "/الترتيب_العام 1 5\n"
        "/ترتيب_نص\n"
        "/احصائيات\n"
        "/احصائيات 1 5\n\n"
        "أوامر الفحص:\n"
        "/الأيام\n"
        "/فحص 5\n"
        "/مشاركين 5\n"
        "/اسطورة 5\n"
        "/مقارنة 4 5\n\n"
        "أوامر الأمان:\n"
        "/مسح_نتائج 5\n"
        "/مسح_يوم 5\n"
        "/مسح_الكل تأكيد\n"
        "/استرجاع_آخر\n"
        "/قفل_يوم 5\n"
        "/فتح_يوم 5\n\n"
        "مثال النتائج:\n"
        "/نتائج 5\n\n"
        "الأهداف:\n"
        "داروين نونيز|3\n\n"
        "الكلين شيت:\n"
        "أوناي سيمون"
    )


async def add_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    day = get_day(text)

    if is_locked(day):
        await update.message.reply_text(f"اليوم {day} مقفل ✅\nلفتحه اكتب: /فتح_يوم {day}")
        return

    lines = text.splitlines()[1:]
    data = {}
    bad_lines = []

    for line in lines:
        line = line.strip()
        if not line:
            continue
        if "|" not in line:
            bad_lines.append(line)
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) != 6:
            bad_lines.append(line)
            continue
        participant = normalize_name(parts[0])
        data[participant] = [normalize_name(x) for x in parts[1:]]

    if not data:
        await update.message.reply_text("ما لقيت مشاركين بصيغة صحيحة.\nالصيغة:\n/اضافه 5\nفهد فارس|الحارس|لاعب1|لاعب2|لاعب3|الكابتن")
        return

    backup_files(f"before_add_day_{day}", files=[excel_file(day), LOCKED_FILE])
    file_name, unknown = update_day_data(day, data)

    caption = f"تم إنشاء/تحديث ملف اليوم {day} ✅\nعدد المشاركين المرسلين: {len(data)}"
    if unknown:
        caption += "\n\n⚠️ أسماء مشاركين غير موجودة بالقائمة:\n" + "\n".join(unknown)
    if bad_lines:
        caption += "\n\n⚠️ أسطر لم أفهمها:\n" + "\n".join(bad_lines[:5])

    with open(file_name, "rb") as file:
        await update.message.reply_document(document=file, filename=file_name, caption=caption)


async def results_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    day = get_day(text)

    if is_locked(day):
        await update.message.reply_text(f"اليوم {day} مقفل ✅\nلفتحه اكتب: /فتح_يوم {day}")
        return

    if not os.path.exists(excel_file(day)):
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}. أضف التشكيلات أولًا.")
        return

    goals_count, clean_sheets = parse_results(text)
    goal_missing, clean_missing = validate_results_names(day, goals_count, clean_sheets)

    backup_files(f"before_results_day_{day}", files=[excel_file(day), LOCKED_FILE])
    file_name = calculate_points(day, goals_count, clean_sheets)

    goals_text = "\n".join([f"- {name}: {count} هدف = {GOAL_POINTS.get(count, count * 5)} نقطة" for name, count in goals_count.items()]) or "لا يوجد"
    clean_text = "\n".join([f"- {name}" for name in clean_sheets]) or "لا يوجد"

    rows = read_day_rows(day)
    max_score = max([r["total"] for r in rows], default=0)
    winners = [r["participant"] for r in rows if r["total"] == max_score and max_score > 0]
    legends_text = "، ".join(winners) if winners else "لا يوجد"

    warnings = []
    if goal_missing:
        warnings.append("⚠️ هدافون غير موجودين في تشكيلات اليوم:\n" + "\n".join(goal_missing))
    if clean_missing:
        warnings.append("⚠️ حراس كلين شيت غير موجودين في تشكيلات اليوم:\n" + "\n".join(clean_missing))

    caption = (
        f"تم حساب نقاط اليوم {day} ✅\n\n"
        f"الأهداف:\n{goals_text}\n\n"
        f"الكلين شيت:\n{clean_text}\n\n"
        f"🏆 أسطورة اليوم: {legends_text} — {max_score} نقطة"
    )
    if warnings:
        caption += "\n\n" + "\n\n".join(warnings)

    with open(file_name, "rb") as file:
        await update.message.reply_document(document=file, filename=file_name, caption=caption)


async def overall(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nums = get_numbers(update.message.text)
    if len(nums) >= 2:
        start_day, end_day = nums[0], nums[1]
    elif len(nums) == 1:
        start_day, end_day = 1, nums[0]
    else:
        start_day, end_day = 1, 31

    file_name, days_found, stats = create_overall_ranking(start_day, end_day)
    if not days_found:
        await update.message.reply_text("ما لقيت أي ملفات أيام محسوبة في النطاق المطلوب.")
        return

    with open(file_name, "rb") as file:
        await update.message.reply_document(
            document=file,
            filename=file_name,
            caption=(
                f"تم إنشاء الترتيب العام ✅\n"
                f"النطاق: من اليوم {start_day} إلى اليوم {end_day}\n"
                f"الأيام المحسوبة: {', '.join(map(str, days_found))}"
            )
        )


async def ranking_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nums = get_numbers(update.message.text)
    if len(nums) >= 2:
        start_day, end_day = nums[0], nums[1]
    elif len(nums) == 1:
        start_day, end_day = 1, nums[0]
    else:
        start_day, end_day = 1, 31
    stats = collect_stats(start_day, end_day)
    await update.message.reply_text(build_ranking_text(stats, start_day, end_day))


async def list_days(update: Update, context: ContextTypes.DEFAULT_TYPE):
    days = get_existing_days(1, 31)
    if not days:
        await update.message.reply_text("ما فيه أيام محفوظة.")
        return

    locked = load_locked_days()
    lines = ["📅 الأيام الموجودة:"]
    for day in days:
        rows = read_day_rows(day)
        participants = sum(1 for r in rows if r["participated"])
        scored = any(r["total"] > 0 for r in rows)
        lock = "🔒" if str(day) in locked else "🔓"
        calc = "محسوب" if scored else "بدون نتائج"
        lines.append(f"اليوم {day} {lock} — {participants} مشارك — {calc}")
    await update.message.reply_text("\n".join(lines))


async def inspect_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    rows = read_day_rows(day)
    if not rows:
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}.")
        return

    participants = [r for r in rows if r["participated"]]
    max_score = max([r["total"] for r in rows], default=0)
    winners = [r["participant"] for r in rows if r["total"] == max_score and max_score > 0]
    zeros = [r["participant"] for r in participants if r["total"] == 0]
    avg = round(sum(r["total"] for r in rows) / len(rows), 2) if rows else 0
    locked = "مقفل 🔒" if is_locked(day) else "مفتوح 🔓"

    lines = [
        f"🔎 فحص اليوم {day}",
        f"الحالة: {locked}",
        f"عدد المشاركين: {len(participants)}",
        f"أعلى نقاط: {max_score}",
        f"أسطورة اليوم: {'، '.join(winners) if winners else 'لا يوجد'}",
        f"عدد اللي جابوا 0: {len(zeros)}",
        f"متوسط النقاط: {avg}",
    ]
    if zeros:
        lines.append("\nالأصفار:")
        lines.extend([f"- {name}" for name in zeros])
    await update.message.reply_text("\n".join(lines))


async def participants_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    rows = read_day_rows(day)
    if not rows:
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}.")
        return

    yes = [r["participant"] for r in rows if r["participated"]]
    no = [r["participant"] for r in rows if not r["participated"]]
    lines = [f"👥 مشاركو اليوم {day}:"]
    lines.extend([f"✅ {name}" for name in yes] or ["لا يوجد"])
    lines.append("\nلم يشارك:")
    lines.extend([f"❌ {name}" for name in no] or ["لا يوجد"])
    await update.message.reply_text("\n".join(lines))


async def legend_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    rows = read_day_rows(day)
    if not rows:
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}.")
        return
    max_score = max([r["total"] for r in rows], default=0)
    winners = [r["participant"] for r in rows if r["total"] == max_score and max_score > 0]
    if not winners:
        await update.message.reply_text(f"ما فيه أسطورة لليوم {day} حتى الآن.")
        return
    await update.message.reply_text(f"🏆 أسطورة اليوم {ordinal_day(day)}:\n" + "\n".join([f"- {w} — {max_score} نقطة" for w in winners]))


async def compare_days(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nums = get_numbers(update.message.text)
    if len(nums) < 2:
        await update.message.reply_text("اكتب رقمين، مثال:\n/مقارنة 4 5\nيعني مقارنة الترتيب بعد اليوم 4 وبعد اليوم 5")
        return

    old_end, new_end = nums[0], nums[1]
    old_stats = collect_stats(1, old_end)
    new_stats = collect_stats(1, new_end)
    if not old_stats["days"] or not new_stats["days"]:
        await update.message.reply_text("ما لقيت أيام كافية للمقارنة.")
        return

    lines = [f"📊 مقارنة الترتيب بعد اليوم {old_end} وبعد اليوم {new_end}:", ""]
    old_pos = {name: i for i, name in enumerate(old_stats["ranking"], start=1)}
    new_pos = {name: i for i, name in enumerate(new_stats["ranking"], start=1)}

    for name in new_stats["ranking"]:
        diff_points = new_stats["totals"][name] - old_stats["totals"].get(name, 0)
        move = old_pos.get(name, new_pos[name]) - new_pos[name]
        move_text = "ثابت"
        if move > 0:
            move_text = f"صعد {move}"
        elif move < 0:
            move_text = f"نزل {abs(move)}"
        lines.append(f"{name}: +{diff_points} نقطة | {move_text}")

    await update.message.reply_text("\n".join(lines))


async def clear_all(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text != "/مسح_الكل تأكيد":
        await update.message.reply_text(
            "⚠️ أمر مسح الكل خطير.\n\n"
            "للتنفيذ اكتب بالضبط:\n"
            "/مسح_الكل تأكيد\n\n"
            "راح أنقل الملفات لنسخة احتياطية بدل حذفها."
        )
        return

    folder, files = backup_files("clear_all", move=True)
    if not files:
        await update.message.reply_text("ما لقيت ملفات أمسحها.")
        return

    await update.message.reply_text(f"✅ تم نقل الملفات للنسخة الاحتياطية.\n📁 {folder}\n📌 عدد الملفات: {len(files)}")


async def clear_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nums = get_numbers(update.message.text)
    if not nums:
        await update.message.reply_text("اكتب رقم اليوم، مثال:\n/مسح_يوم 5")
        return
    day = str(nums[0])
    if is_locked(day):
        await update.message.reply_text(f"اليوم {day} مقفل ✅\nلفتحه اكتب: /فتح_يوم {day}")
        return

    file_name = excel_file(day)
    if not os.path.exists(file_name):
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}.")
        return

    folder, _ = backup_files(f"before_clear_day_{day}", files=[file_name, LOCKED_FILE])
    os.remove(file_name)
    await update.message.reply_text(f"تم مسح ملف اليوم {day} ✅\nتم حفظ نسخة احتياطية: {folder}")


async def clear_results(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    if is_locked(day):
        await update.message.reply_text(f"اليوم {day} مقفل ✅\nلفتحه اكتب: /فتح_يوم {day}")
        return

    backup_files(f"before_clear_results_{day}", files=[excel_file(day), LOCKED_FILE])
    file_name = clear_day_points(day)
    if not file_name:
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}.")
        return

    with open(file_name, "rb") as file:
        await update.message.reply_document(document=file, filename=file_name, caption=f"تم مسح نتائج اليوم {day} مع بقاء المشاركين ✅")


async def restore_last(update: Update, context: ContextTypes.DEFAULT_TYPE):
    folder = latest_backup_folder()
    if not folder:
        await update.message.reply_text("ما فيه نسخة احتياطية أسترجعها.")
        return

    # نسخة احتياطية قبل الاسترجاع
    backup_files("before_restore")

    restored = []
    for filename in os.listdir(folder):
        if filename.endswith(".xlsx") or filename == LOCKED_FILE:
            shutil.copy2(os.path.join(folder, filename), filename)
            restored.append(filename)

    if not restored:
        await update.message.reply_text(f"النسخة {folder} ما فيها ملفات قابلة للاسترجاع.")
        return

    await update.message.reply_text(f"✅ تم استرجاع آخر نسخة احتياطية:\n📁 {folder}\n\nالملفات:\n" + "\n".join(restored))


async def lock_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    days = load_locked_days()
    days.add(str(day))
    save_locked_days(days)
    await update.message.reply_text(f"تم قفل اليوم {day} 🔒")


async def unlock_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    days = load_locked_days()
    days.discard(str(day))
    save_locked_days(days)
    await update.message.reply_text(f"تم فتح اليوم {day} 🔓")


async def dashboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        nums = get_numbers(update.message.text)

        if len(nums) >= 2:
            start_day, end_day = nums[0], nums[1]
        elif len(nums) == 1:
            start_day, end_day = 1, nums[0]
        else:
            start_day, end_day = 1, 31

        await update.message.reply_text("جاري إنشاء ملف الإحصائيات... ⏳")

        file_name, stats = create_dashboard(start_day, end_day)

        if not stats.get("days"):
            await update.message.reply_text(
                "ما لقيت أيام لإحصائيات الداشبورد.\n"
                "تأكد أن ملفات الأيام موجودة مثل:\n"
                "fantasy_day_1.xlsx\n"
                "fantasy_day_2.xlsx"
            )
            return

        if not os.path.exists(file_name):
            await update.message.reply_text(
                "صار خطأ: ملف الإحصائيات ما انحفظ.\n"
                f"اسم الملف المتوقع: {file_name}"
            )
            return

        caption = (
            "تم إنشاء ملف الإحصائيات الكامل ✅\n"
            f"النطاق: من اليوم {start_day} إلى اليوم {end_day}\n"
            f"الأيام المحسوبة: {', '.join(map(str, stats['days']))}\n\n"
            "الصفحات:\n"
            "لوحة عامة\n"
            "الترتيب العام\n"
            "تطور النقاط\n"
            "تحليل الأيام\n"
            "تحليل المشاركين\n"
            "تحليل الكباتن\n"
            "تحليل الحراس\n"
            "تحليل اللاعبين"
        )

        with open(file_name, "rb") as file:
            await update.message.reply_document(
                document=file,
                filename=file_name,
                caption=caption
            )

    except Exception as e:
        await update.message.reply_text(
            "صار خطأ أثناء إنشاء الإحصائيات ❌\n\n"
            f"السبب:\n{e}"
        )


def main():
    if not TOKEN:
        raise RuntimeError("ضع توكن البوت في متغير البيئة BOT_TOKEN")

    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/start"), start))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اضافه"), add_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج"), results_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الترتيب_العام"), overall))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/ترتيب_نص"), ranking_text))

    # الإحصائيات — يدعم:
    # /احصائيات
    # /احصائيات 1 5
    # احصائيات
    # احصائيات 1 5
    app.add_handler(MessageHandler(
        filters.TEXT & filters.Regex(r"^/?(?:احصائيات|إحصائيات)(?:\s|$)"),
        dashboard
    ))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/(الأيام|الايام)"), list_days))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/فحص"), inspect_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مشاركين"), participants_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اسطورة"), legend_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مقارنة"), compare_days))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_الكل"), clear_all))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_يوم"), clear_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_نتائج"), clear_results))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استرجاع_آخر"), restore_last))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/قفل_يوم"), lock_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/فتح_يوم"), unlock_day))

    app.run_polling()


if __name__ == "__main__":
    main()