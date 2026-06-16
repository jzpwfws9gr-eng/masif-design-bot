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
    ws.sheet_view.rightToLeft = True
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 22
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
    stats = collect_stats(start_day, end_day)
    days = stats["days"]
    file_name = "fantasy_dashboard.xlsx"

    wb = Workbook()
    wb.remove(wb.active)

    # 1) لوحة عامة
    ws = wb.create_sheet("لوحة عامة")
    ws.sheet_view.rightToLeft = True
    ws["A1"] = "🏆 لوحة إحصائيات فانتزي المصيف 2026"
    ws.merge_cells("A1:D1")
    style_title_cell(ws["A1"])

    leader = stats["ranking"][0] if stats["ranking"] else "-"
    runner = stats["ranking"][1] if len(stats["ranking"]) > 1 else None
    gap = stats["totals"][leader] - stats["totals"][runner] if runner else 0
    most_legend = max(PARTICIPANTS, key=lambda n: stats["daily_wins"][n]) if PARTICIPANTS else "-"
    highest_daily = 0
    highest_daily_name = "-"
    highest_daily_day = "-"
    for day, info in stats["per_day"].items():
        if info["max_score"] > highest_daily:
            highest_daily = info["max_score"]
            highest_daily_name = "، ".join(info["winners"])
            highest_daily_day = day

    summary = [
        ["إجمالي المشاركين", len(PARTICIPANTS)],
        ["عدد الأيام المحسوبة", len(days)],
        ["الأيام", ", ".join(map(str, days)) if days else "لا يوجد"],
        ["المتصدر", leader],
        ["نقاط المتصدر", stats["totals"].get(leader, 0)],
        ["الفرق بين الأول والثاني", gap],
        ["أكثر أسطورة يوم", f"{most_legend} ({stats['daily_wins'][most_legend]})"],
        ["أعلى نقاط يومية", f"{highest_daily_name} - {highest_daily} نقطة - اليوم {highest_daily_day}"],
    ]
    ws.append([])
    for row in summary:
        ws.append(row)

    start = 12
    ws.cell(row=start, column=1).value = "أفضل 5 مشاركين"
    ws.cell(row=start, column=4).value = "أقل 5 مشاركين"
    for cell in (ws.cell(row=start, column=1), ws.cell(row=start, column=4)):
        cell.font = Font(bold=True, size=14, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")

    ws.cell(row=start + 1, column=1).value = "المشارك"
    ws.cell(row=start + 1, column=2).value = "النقاط"
    ws.cell(row=start + 1, column=4).value = "المشارك"
    ws.cell(row=start + 1, column=5).value = "النقاط"

    top5 = stats["ranking"][:5]
    bottom5 = list(reversed(stats["ranking"][-5:]))
    for i in range(5):
        if i < len(top5):
            ws.cell(row=start + 2 + i, column=1).value = top5[i]
            ws.cell(row=start + 2 + i, column=2).value = stats["totals"][top5[i]]
        if i < len(bottom5):
            ws.cell(row=start + 2 + i, column=4).value = bottom5[i]
            ws.cell(row=start + 2 + i, column=5).value = stats["totals"][bottom5[i]]

    # جدول مختصر للترتيب للرسم
    chart_start = 20
    ws.cell(row=chart_start, column=1).value = "المشارك"
    ws.cell(row=chart_start, column=2).value = "المجموع"
    for i, name in enumerate(stats["ranking"], start=chart_start + 1):
        ws.cell(row=i, column=1).value = name
        ws.cell(row=i, column=2).value = stats["totals"][name]
    add_bar_chart(ws, "الترتيب العام", 2, 2, chart_start, chart_start + len(stats["ranking"]), 1, chart_start + 1, chart_start + len(stats["ranking"]), "G3")
    style_dashboard_sheet(ws)

    # 2) الترتيب العام
    ws = wb.create_sheet("الترتيب العام")
    ws.append(["المركز", "المشارك", "المجموع", "أسطورة اليوم", "عدد المشاركات", "متوسط النقاط"])
    current_rank = 0
    last_score = None
    real_index = 0
    for name in stats["ranking"]:
        real_index += 1
        score = stats["totals"][name]
        if score != last_score:
            current_rank = real_index
            last_score = score
        pc = stats["participation_count"][name]
        avg = round(score / pc, 2) if pc else 0
        ws.append([current_rank, name, score, stats["daily_wins"][name], pc, avg])
    style_sheet(ws)
    add_bar_chart(ws, "الترتيب العام بالنقاط", 3, 3, 1, ws.max_row, 2, 2, ws.max_row, "H2")

    # 3) تطور النقاط
    ws = wb.create_sheet("تطور النقاط")
    ws.append(["اليوم"] + stats["ranking"])
    for day in days:
        ws.append([day] + [stats["cumulative_by_day"][name].get(day, 0) for name in stats["ranking"]])
    style_sheet(ws)
    if days:
        add_line_chart(ws, "تطور مجموع النقاط يومًا بعد يوم", 2, 1 + len(stats["ranking"]), 1, 2, ws.max_row, 1, "A16")

    # 4) تحليل الأيام
    ws = wb.create_sheet("تحليل الأيام")
    ws.append(["اليوم", "أسطورة اليوم", "أعلى نقاط", "عدد المشاركين", "عدد الأصفار", "متوسط النقاط", "أقل نقاط"])
    for day in days:
        info = stats["per_day"].get(day, {})
        ws.append([
            day,
            "، ".join(info.get("winners", [])) or "-",
            info.get("max_score", 0),
            info.get("participants", 0),
            info.get("zeros", 0),
            info.get("avg", 0),
            info.get("min_score", 0),
        ])
    style_sheet(ws)
    add_column_chart(ws, "أعلى نقاط كل يوم", 3, 1, 2, ws.max_row, 1, "I2")
    add_column_chart(ws, "متوسط نقاط كل يوم", 6, 1, 2, ws.max_row, 1, "I18")

    # 5) تحليل المشاركين
    ws = wb.create_sheet("تحليل المشاركين")
    ws.append(["المشارك", "المجموع", "عدد المشاركات", "نسبة المشاركة", "أسطورة اليوم", "أفضل يوم", "نقاط أفضل يوم", "أسوأ يوم", "نقاط أسوأ يوم", "أيام صفر"])
    total_days = len(days)
    for name in stats["ranking"]:
        day_scores = stats["scores_by_day"][name]
        if day_scores:
            best_day = max(day_scores, key=lambda d: day_scores[d])
            worst_day = min(day_scores, key=lambda d: day_scores[d])
            best_score = day_scores[best_day]
            worst_score = day_scores[worst_day]
        else:
            best_day = worst_day = "-"
            best_score = worst_score = 0
        pc = stats["participation_count"][name]
        pct = f"{round((pc / total_days) * 100, 1)}%" if total_days else "0%"
        ws.append([name, stats["totals"][name], pc, pct, stats["daily_wins"][name], best_day, best_score, worst_day, worst_score, stats["zero_days"][name]])
    style_sheet(ws)
    add_bar_chart(ws, "مجموع نقاط المشاركين", 2, 2, 1, ws.max_row, 1, 2, ws.max_row, "L2")

    # 6) تحليل الكباتن
    ws = wb.create_sheet("تحليل الكباتن")
    ws.append(["المشارك", "نقاط الكابتن", "عدد المشاركات", "متوسط نقاط الكابتن"])
    for name in stats["ranking"]:
        pc = stats["participation_count"][name]
        cap_pts = stats["captain_points_by_participant"][name]
        ws.append([name, cap_pts, pc, round(cap_pts / pc, 2) if pc else 0])

    start2 = ws.max_row + 3
    ws.cell(row=start2, column=1).value = "اللاعب"
    ws.cell(row=start2, column=2).value = "مرات اختياره كابتن"
    ws.cell(row=start2, column=3).value = "نقاطه ككابتن"
    row_i = start2 + 1
    for player, count in stats["captain_choice_count"].most_common():
        ws.cell(row=row_i, column=1).value = player
        ws.cell(row=row_i, column=2).value = count
        ws.cell(row=row_i, column=3).value = stats["captain_points_by_player"][player]
        row_i += 1
    style_sheet(ws)
    add_bar_chart(ws, "أكثر مشارك استفاد من الكابتن", 2, 2, 1, 1 + len(PARTICIPANTS), 1, 2, 1 + len(PARTICIPANTS), "F2")
    if row_i > start2 + 1:
        add_column_chart(ws, "أكثر لاعب تم اختياره كابتن", 2, start2, start2 + 1, row_i - 1, 1, "F20")

    # 7) تحليل الحراس
    ws = wb.create_sheet("تحليل الحراس")
    ws.append(["المشارك", "نقاط الحراس", "نجاح الكلين شيت", "عدد المشاركات", "نسبة نجاح الحارس"])
    for name in stats["ranking"]:
        pc = stats["participation_count"][name]
        success = stats["keeper_success_by_participant"][name]
        ws.append([name, stats["keeper_points_by_participant"][name], success, pc, f"{round((success / pc) * 100, 1)}%" if pc else "0%"])

    start2 = ws.max_row + 3
    ws.cell(row=start2, column=1).value = "الحارس"
    ws.cell(row=start2, column=2).value = "مرات اختياره"
    ws.cell(row=start2, column=3).value = "نقاطه"
    row_i = start2 + 1
    for keeper, count in stats["keeper_choice_count"].most_common():
        ws.cell(row=row_i, column=1).value = keeper
        ws.cell(row=row_i, column=2).value = count
        ws.cell(row=row_i, column=3).value = stats["keeper_points_by_keeper"][keeper]
        row_i += 1
    style_sheet(ws)
    add_bar_chart(ws, "أكثر مشارك استفاد من الحارس", 2, 2, 1, 1 + len(PARTICIPANTS), 1, 2, 1 + len(PARTICIPANTS), "G2")
    if row_i > start2 + 1:
        add_column_chart(ws, "أكثر حارس تم اختياره", 2, start2, start2 + 1, row_i - 1, 1, "G20")

    # 8) تحليل اللاعبين
    ws = wb.create_sheet("تحليل اللاعبين")
    ws.append(["اللاعب", "مرات الاختيار", "نقاط اللاعب", "مرات كابتن", "متوسط نقاط لكل اختيار", "مرات اختير ولم يسجل"])
    all_players = set(stats["player_choice_count"].keys()) | set(stats["captain_selections_by_player"].keys())
    sorted_players = sorted(all_players, key=lambda p: (stats["player_choice_count"][p], stats["player_points"][p]), reverse=True)
    for player in sorted_players:
        count = stats["player_choice_count"][player]
        pts = stats["player_points"][player]
        ws.append([player, count, pts, stats["captain_selections_by_player"][player], round(pts / count, 2) if count else 0, stats["player_zero_selections"][player]])
    style_sheet(ws)
    if ws.max_row >= 2:
        add_column_chart(ws, "أكثر اللاعبين اختيارًا", 2, 1, 2, min(ws.max_row, 16), 1, "H2")
        add_column_chart(ws, "أكثر لاعب نافع بالنقاط", 3, 1, 2, min(ws.max_row, 16), 1, "H18")

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