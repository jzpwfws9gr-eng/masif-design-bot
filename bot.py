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
    def sort_key(x):
        try:
            return int(x)
        except Exception:
            return 0
    with open(LOCKED_FILE, "w", encoding="utf-8") as f:
        json.dump({"locked_days": sorted([str(x) for x in days], key=sort_key)}, f, ensure_ascii=False, indent=2)


def is_locked(day):
    return str(day) in load_locked_days()


def current_data_files():
    files = []
    for filename in os.listdir("."):
        if (
            filename.startswith("fantasy_day_") and filename.endswith(".xlsx")
        ) or filename in ("overall_ranking.xlsx", "fantasy_dashboard.xlsx", LOCKED_FILE):
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

    for idx, header in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=idx).value = header

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

def create_dashboard(start_day=1, end_day=31):
    from collections import defaultdict, Counter
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    try:
        from openpyxl.chart import DoughnutChart
    except Exception:
        DoughnutChart = PieChart

    stats = collect_stats(start_day, end_day)
    days = stats["days"]
    file_name = "fantasy_dashboard.xlsx"

    wb = Workbook()
    wb.remove(wb.active)

    # ألوان Dashboard فاتح مثل الصورة الثانية
    BG = "F5F7FB"
    WHITE = "FFFFFF"
    NAVY = "243B53"
    BLUE = "4F7DF3"
    PURPLE = "7C3AED"
    L_PURPLE = "EFE7FF"
    L_BLUE = "EAF1FF"
    GREEN = "12B886"
    L_GREEN = "E7F8F1"
    GOLD = "F2B705"
    L_GOLD = "FFF6D6"
    RED = "F06565"
    L_RED = "FFECEC"
    TEXT = "1F2937"
    MUTED = "6B7280"
    GRID = "D9E2EF"
    HEADER = "3B4A6B"

    thin = Side(style="thin", color=GRID)
    no_side = Side(style=None)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    INVALID_SELECTIONS = {"", "لم يشارك", "تم حجب المشاركة / تأخير"}

    def is_real_selection(value):
        return normalize_name(value) not in INVALID_SELECTIONS

    def points_source(base_pts, cap_extra):
        parts = []
        try:
            base_pts = int(base_pts or 0)
        except Exception:
            base_pts = 0
        try:
            cap_extra = int(cap_extra or 0)
        except Exception:
            cap_extra = 0

        if base_pts > 0:
            goals = base_pts // 5
            if goals <= 1:
                parts.append("هدف")
            elif goals == 2:
                parts.append("هدفين")
            elif goals <= 10:
                parts.append(f"{goals} أهداف")
            else:
                parts.append(f"{base_pts} نقطة")
        if cap_extra > 0:
            parts.append("كابتن")
        return " + ".join(parts) if parts else "بدون نقاط"

    def safe_pct(part, total):
        return f"{round((part / total) * 100, 1)}%" if total else "0%"

    def safe_avg(total, count):
        return round(total / count, 2) if count else 0

    def sheet_setup(ws, title=None):
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False
        for col in range(1, 22):
            ws.column_dimensions[get_column_letter(col)].width = 16
        for row in range(1, 90):
            ws.row_dimensions[row].height = 22
        if title:
            ws["A1"] = title
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
            c = ws["A1"]
            c.font = Font(bold=True, size=18, color=WHITE)
            c.fill = PatternFill("solid", fgColor=HEADER)
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 34

    def style_table(ws, min_row, max_row, min_col, max_col, header_row=None, autofilter=True):
        header_row = header_row or min_row
        if max_row < min_row or max_col < min_col:
            return
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if isinstance(cell.value, str):
                    cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
                if r == header_row:
                    cell.fill = PatternFill("solid", fgColor=HEADER)
                    cell.font = Font(bold=True, size=11, color=WHITE)
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    cell.font = Font(size=11, color=TEXT)
                    cell.fill = PatternFill("solid", fgColor=("F8FAFC" if r % 2 == 0 else WHITE))
        ws.row_dimensions[header_row].height = 30
        if autofilter:
            ws.auto_filter.ref = f"{get_column_letter(min_col)}{header_row}:{get_column_letter(max_col)}{max_row}"

    def section_title(ws, row, col, title, width=4, color=HEADER):
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + width - 1)
        c = ws.cell(row=row, column=col)
        c.value = title
        c.font = Font(bold=True, size=13, color=WHITE)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
        ws.row_dimensions[row].height = 30

    def card(ws, cell_range, title, value, color=BLUE, sub_text=None):
        ws.merge_cells(cell_range)
        c = ws[cell_range.split(":")[0]]
        text = f"{title}\n{value}"
        if sub_text:
            text += f"\n{sub_text}"
        c.value = text
        c.font = Font(bold=True, size=13, color=WHITE)
        c.fill = PatternFill("solid", fgColor=color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def light_card(ws, cell_range, title, value, color_fill=L_BLUE, color_text=TEXT):
        ws.merge_cells(cell_range)
        c = ws[cell_range.split(":")[0]]
        c.value = f"{title}\n{value}"
        c.font = Font(bold=True, size=12, color=color_text)
        c.fill = PatternFill("solid", fgColor=color_fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    def add_bar_chart(ws, title, data_col, header_row, min_row, max_row, cats_col, anchor, width=12, height=7, chart_type="bar"):
        if max_row < min_row:
            return
        chart = BarChart()
        chart.type = chart_type
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

    def add_line_chart_local(ws, title, min_col, max_col, header_row, min_row, max_row, cats_col, anchor, width=18, height=8):
        if max_row < min_row or max_col < min_col:
            return
        chart = LineChart()
        chart.title = title
        chart.y_axis.title = "النقاط"
        chart.x_axis.title = "الجولة"
        data = Reference(ws, min_col=min_col, max_col=max_col, min_row=header_row, max_row=max_row)
        cats = Reference(ws, min_col=cats_col, min_row=min_row, max_row=max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = width
        chart.height = height
        ws.add_chart(chart, anchor)

    def add_donut_chart(ws, title, labels_col, values_col, min_row, max_row, anchor, width=9, height=7):
        if max_row < min_row:
            return
        chart = DoughnutChart()
        chart.title = title
        data = Reference(ws, min_col=values_col, min_row=min_row, max_row=max_row)
        labels = Reference(ws, min_col=labels_col, min_row=min_row, max_row=max_row)
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(labels)
        chart.width = width
        chart.height = height
        if hasattr(chart, "holeSize"):
            chart.holeSize = 55
        ws.add_chart(chart, anchor)

    # -------------------- حسابات إضافية للصفحات الجديدة --------------------
    participant_count = len(PARTICIPANTS)
    total_slots = participant_count * len(days) if days else 0
    total_active_slots = 0

    day_summary = {}
    active_scores_by_participant = {name: {} for name in PARTICIPANTS}

    keeper_participants = defaultdict(set)
    keeper_rounds = defaultdict(set)
    keeper_clean_rounds = defaultdict(set)
    keeper_round_detail = defaultdict(lambda: {"selected": 0, "people": set(), "clean": False, "total_points": 0})

    captain_by_participant = {name: {"count": 0, "success": 0, "points": 0, "best_player": "-", "best_day": "-", "best_points": -1} for name in PARTICIPANTS}
    captain_player_count = Counter()
    captain_player_points = Counter()

    player_participants = defaultdict(set)
    player_round_detail = defaultdict(lambda: {"selected": 0, "captains": 0, "base_points": 0, "distributed": 0})
    player_actual_by_day = defaultdict(dict)
    player_total_distributed = Counter()
    player_actual_total = Counter()
    player_choice_count = Counter()
    player_captain_count = Counter()

    selection_detail_rows = []

    for day in days:
        rows = read_day_rows(day)
        active_totals = []
        total_day_points = 0

        for row in rows:
            name = row["participant"]
            total = row["total"]
            total_day_points += total

            if row["participated"]:
                total_active_slots += 1
                active_totals.append(total)
                active_scores_by_participant[name][day] = total

                # الحراس
                keeper = row["keeper"]
                if is_real_selection(keeper):
                    keeper_participants[keeper].add(name)
                    keeper_rounds[keeper].add(day)
                    kd = keeper_round_detail[(day, keeper)]
                    kd["selected"] += 1
                    kd["people"].add(name)
                    kd["total_points"] += row["keeper_points"]
                    if row["keeper_points"] > 0:
                        kd["clean"] = True
                        keeper_clean_rounds[keeper].add(day)


                # الكباتن - حسب المشارك
                captain = row["captain"]
                if is_real_selection(captain):
                    captain_by_participant[name]["count"] += 1
                    captain_by_participant[name]["points"] += row["captain_points"]
                    captain_player_count[captain] += 1
                    captain_player_points[captain] += row["captain_points"]
                    if row["captain_points"] > 0:
                        captain_by_participant[name]["success"] += 1
                    if row["captain_points"] > captain_by_participant[name]["best_points"]:
                        captain_by_participant[name]["best_points"] = row["captain_points"]
                        captain_by_participant[name]["best_player"] = captain
                        captain_by_participant[name]["best_day"] = day

                # اللاعبون
                for p_key, pts_key in (("p1", "p1_points"), ("p2", "p2_points"), ("p3", "p3_points")):
                    player = row[p_key]
                    base_pts = row[pts_key]
                    if is_real_selection(player):
                        player_participants[player].add(name)
                        player_choice_count[player] += 1
                        pd = player_round_detail[(day, player)]
                        pd["selected"] += 1
                        pd["base_points"] = max(pd["base_points"], base_pts)
                        player_actual_by_day[player][day] = max(player_actual_by_day[player].get(day, 0), base_pts)

                        cap_extra = row["captain_points"] if row["captain"] == player else 0
                        if row["captain"] == player:
                            pd["captains"] += 1
                            player_captain_count[player] += 1

                        selection_detail_rows.append([
                            day,
                            name,
                            player,
                            "نعم 👑" if row["captain"] == player else "لا",
                            base_pts,
                            cap_extra,
                            base_pts + cap_extra,
                            points_source(base_pts, cap_extra),
                        ])

        for (d, player), detail in list(player_round_detail.items()):
            if d == day:
                detail["distributed"] = detail["base_points"] * detail["selected"] + detail["base_points"] * detail["captains"]

        info = stats["per_day"].get(day, {})
        winners = info.get("winners", [])
        day_summary[day] = {
            "participants": info.get("participants", 0),
            "non_participants": max(participant_count - info.get("participants", 0), 0),
            "total_points": total_day_points,
            "avg_active": round(sum(active_totals) / len(active_totals), 2) if active_totals else 0,
            "max_score": info.get("max_score", 0),
            "winners": winners,
            "legends_count": len(winners),
        }

    for player, by_day in player_actual_by_day.items():
        player_actual_total[player] = sum(by_day.values())

    for (day, player), detail in player_round_detail.items():
        player_total_distributed[player] += detail["distributed"]

    ranking = stats["ranking"]
    leader = ranking[0] if ranking else "-"
    leader_points = stats["totals"].get(leader, 0)

    highest_daily = 0
    highest_daily_name = "-"
    highest_daily_day = "-"
    for day, info in day_summary.items():
        if info["max_score"] > highest_daily:
            highest_daily = info["max_score"]
            highest_daily_name = "، ".join(info["winners"]) or "-"
            highest_daily_day = day

    most_legend = "-"
    if PARTICIPANTS:
        most_legend = max(PARTICIPANTS, key=lambda n: stats["daily_wins"][n])
        if stats["daily_wins"].get(most_legend, 0) == 0:
            most_legend = "-"

    best_keeper = "-"
    if keeper_rounds:
        best_keeper = max(keeper_rounds.keys(), key=lambda k: stats["keeper_points_by_keeper"][k])

    best_captain_participant = "-"
    if ranking:
        best_captain_participant = max(ranking, key=lambda n: captain_by_participant[n]["points"])

    # أكثر لاعب وزع نقاط على المشاركين في جولة واحدة
    top_round_player = None
    if player_round_detail:
        top_round_player = max(player_round_detail.items(), key=lambda x: x[1]["distributed"])

    best_actual_player = "-"
    if player_actual_total:
        best_actual_player = max(player_actual_total.keys(), key=lambda p: player_actual_total[p])

    most_selected_player = "-"
    if player_choice_count:
        most_selected_player = player_choice_count.most_common(1)[0][0]

    most_captained_player = "-"
    if player_captain_count:
        most_captained_player = player_captain_count.most_common(1)[0][0]

    # اللاعب الذهبي: أكثر لاعب أفاد المشاركين في البطولة كاملة
    golden_player = "-"
    if player_total_distributed:
        golden_player = max(player_total_distributed.keys(), key=lambda p: player_total_distributed[p])

    def build_rank_map_from_scores(score_map):
        ordered = sorted(score_map.keys(), key=lambda n: score_map.get(n, 0), reverse=True)
        ranks = {}
        current_rank = 0
        last_score = None
        real_index = 0
        for n in ordered:
            real_index += 1
            s = score_map.get(n, 0)
            if s != last_score:
                current_rank = real_index
                last_score = s
            ranks[n] = current_rank
        return ranks

    def movement_text(name, current_rank, previous_rank):
        if previous_rank is None:
            return "جديد"
        diff = previous_rank - current_rank
        if diff > 0:
            return f"↑ صعد {diff} مركز" if diff == 1 else f"↑ صعد {diff} مراكز"
        if diff < 0:
            drop = abs(diff)
            return f"↓ نزل مركز" if drop == 1 else f"↓ نزل {drop} مراكز"
        return "— ثابت"

    current_rank_map = build_rank_map_from_scores(stats["totals"])
    previous_rank_map = {}
    if len(days) >= 2:
        prev_day = days[-2]
        prev_scores = {name: stats["cumulative_by_day"][name].get(prev_day, 0) for name in PARTICIPANTS}
        previous_rank_map = build_rank_map_from_scores(prev_scores)

    # -------------------- 1) لوحة عامة Dashboard --------------------
    ws = wb.create_sheet("لوحة عامة")
    sheet_setup(ws)
    ws.sheet_properties.tabColor = PURPLE

    # خلفية/عنوان
    for row in range(1, 50):
        for col in range(1, 14):
            ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=BG)

    ws.merge_cells("A1:M2")
    ws["A1"] = "لوحة فانتزي المصيف 2026"
    ws["A1"].font = Font(bold=True, size=22, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=PURPLE)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 18

    # كروت عليا
    card(ws, "A4:C6", "🏆 المتصدر", f"{leader}\n{leader_points} نقطة", PURPLE)
    card(ws, "D4:F6", "👥 المشاركون", f"{participant_count}", BLUE)
    card(ws, "G4:I6", "📅 الجولات", f"{len(days)}", GREEN)
    card(ws, "J4:M6", "🔥 أعلى نتيجة يومية", f"{highest_daily_name}\n{highest_daily} نقطة - اليوم {highest_daily_day}", GOLD)

    light_card(ws, "A8:C10", "⭐ أكثر أسطورة يوم", f"{most_legend}\n{stats['daily_wins'].get(most_legend, 0)} مرات", L_GOLD)
    light_card(ws, "D8:F10", "🧤 أفضل حارس", f"{best_keeper}\n{stats['keeper_points_by_keeper'].get(best_keeper, 0)} نقطة", L_BLUE)
    light_card(ws, "G8:I10", "👑 أفضل مشارك بالكباتن", f"{best_captain_participant}\n{captain_by_participant.get(best_captain_participant, {}).get('points', 0)} نقطة إضافية", L_PURPLE)
    if top_round_player:
        (top_day, top_player), top_info = top_round_player
        light_card(ws, "J8:M10", "🔥 أكثر لاعب وزّع نقاط على المشاركين في جولة واحدة", f"{top_player} - اليوم {top_day}\nوزّع {top_info['distributed']} نقطة", L_RED)
    else:
        light_card(ws, "J8:M10", "🔥 أكثر لاعب وزّع نقاط على المشاركين في جولة واحدة", "-", L_RED)

    # بيانات مخفية للرسوم
    data_col = 16  # P
    ws.cell(row=1, column=data_col).value = "بيانات"
    # نسبة المشاركة
    ws.cell(row=2, column=data_col).value = "شاركوا"
    ws.cell(row=2, column=data_col + 1).value = total_active_slots
    ws.cell(row=3, column=data_col).value = "لم يشاركوا"
    ws.cell(row=3, column=data_col + 1).value = max(total_slots - total_active_slots, 0)

    # توزيع الأساطير
    leg_start = 6
    ws.cell(row=leg_start, column=data_col).value = "المشارك"
    ws.cell(row=leg_start, column=data_col + 1).value = "مرات أسطورة اليوم"
    leg_rows = [(n, stats["daily_wins"][n]) for n in ranking if stats["daily_wins"][n] > 0][:8]
    if not leg_rows:
        leg_rows = [("-", 0)]
    for i, (name, val) in enumerate(leg_rows, start=leg_start + 1):
        ws.cell(row=i, column=data_col).value = name
        ws.cell(row=i, column=data_col + 1).value = val

    # أفضل 5
    top_start = 17
    ws.cell(row=top_start, column=data_col).value = "المشارك"
    ws.cell(row=top_start, column=data_col + 1).value = "النقاط"
    for i, name in enumerate(ranking[:5], start=top_start + 1):
        ws.cell(row=i, column=data_col).value = name
        ws.cell(row=i, column=data_col + 1).value = stats["totals"][name]

    # نقاط الجولات
    day_start = 26
    ws.cell(row=day_start, column=data_col).value = "الجولة"
    ws.cell(row=day_start, column=data_col + 1).value = "مجموع نقاط الجولة"
    for i, day in enumerate(days, start=day_start + 1):
        ws.cell(row=i, column=data_col).value = day
        ws.cell(row=i, column=data_col + 1).value = day_summary[day]["total_points"]

    # تطور أول 5
    evo_start = 35
    ws.cell(row=evo_start, column=data_col).value = "الجولة"
    for idx, name in enumerate(ranking[:5], start=data_col + 1):
        ws.cell(row=evo_start, column=idx).value = name
    for r, day in enumerate(days, start=evo_start + 1):
        ws.cell(row=r, column=data_col).value = day
        for c, name in enumerate(ranking[:5], start=data_col + 1):
            ws.cell(row=r, column=c).value = stats["cumulative_by_day"][name].get(day, 0)

    # الترتيب العام داخل لوحة عامة
    section_title(ws, 12, 1, "الترتيب العام", 6, HEADER)
    ranking_headers = ["المركز", "المشارك", "النقاط", "الفارق عن المتصدر", "حركة الترتيب", "أسطورة اليوم"]
    for col, h in enumerate(ranking_headers, start=1):
        ws.cell(row=13, column=col).value = h

    current_rank = 0
    last_score = None
    real_index = 0
    rank_row = 14
    for name in stats["ranking"]:
        real_index += 1
        score = stats["totals"][name]
        if score != last_score:
            current_rank = real_index
            last_score = score
        prev_rank = previous_rank_map.get(name) if previous_rank_map else None
        move = movement_text(name, current_rank, prev_rank)
        values = [current_rank, name, score, leader_points - score, move, stats["daily_wins"][name]]
        for col, val in enumerate(values, start=1):
            ws.cell(row=rank_row, column=col).value = val
        # تلوين حركة الترتيب فقط بدون إزعاج الجدول
        move_cell = ws.cell(row=rank_row, column=5)
        if isinstance(move, str) and move.startswith("↑"):
            move_cell.font = Font(bold=True, color=GREEN)
        elif isinstance(move, str) and move.startswith("↓"):
            move_cell.font = Font(bold=True, color=RED)
        elif move == "جديد":
            move_cell.font = Font(bold=True, color=BLUE)
        rank_row += 1

    style_table(ws, 13, max(13, rank_row - 1), 1, 6, header_row=13)
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["E"].width = 20

    # الرسوم تحت جدول الترتيب
    chart_top = max(31, rank_row + 3)
    add_donut_chart(ws, "نسبة المشاركة", data_col, data_col + 1, 2, 3, f"A{chart_top}", width=8, height=7)
    add_donut_chart(ws, "توزيع أسطورة اليوم", data_col, data_col + 1, leg_start + 1, leg_start + len(leg_rows), f"D{chart_top}", width=8, height=7)
    add_bar_chart(ws, "أفضل 5 مشاركين", data_col + 1, top_start, top_start + 1, top_start + len(ranking[:5]), data_col, f"G{chart_top}", width=8, height=7, chart_type="bar")
    add_bar_chart(ws, "مجموع نقاط الجولات", data_col + 1, day_start, day_start + 1, day_start + len(days), data_col, f"J{chart_top}", width=8, height=7, chart_type="col")
    add_line_chart_local(ws, "تطور نقاط أفضل 5", data_col + 1, data_col + min(5, len(ranking)), evo_start, evo_start + 1, evo_start + len(days), data_col, f"A{chart_top + 14}", width=18, height=8)

    for col in range(data_col, data_col + 8):
        ws.column_dimensions[get_column_letter(col)].hidden = True

    # -------------------- 2) تطور النقاط ✅ كما هو --------------------
    ws = wb.create_sheet("تطور النقاط")
    ws.append(["اليوم"] + stats["ranking"])

    for day in days:
        ws.append([day] + [stats["cumulative_by_day"][name].get(day, 0) for name in stats["ranking"]])

    style_sheet(ws)

    if days:
        add_line_chart_local(ws, "تطور مجموع النقاط يومًا بعد يوم", 2, 1 + len(stats["ranking"]), 1, 2, ws.max_row, 1, "A16", width=20, height=10)

    # -------------------- 3) تحليل الأيام ✅ مع تنسيق أسطورة اليوم --------------------
    ws = wb.create_sheet("تحليل الأيام")
    sheet_setup(ws, "تحليل الأيام")
    headers = ["اليوم", "عدد المشاركين", "غير المشاركين", "مجموع نقاط اليوم", "متوسط نقاط المشاركين", "أعلى نقاط", "عدد الأساطير", "أسطورة اليوم"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=2, column=col).value = h

    row_i = 3
    for day in days:
        info = day_summary[day]
        winners_text = "\n".join([f"⭐ {w}" for w in info["winners"]]) if info["winners"] else "-"
        values = [
            day,
            info["participants"],
            info["non_participants"],
            info["total_points"],
            info["avg_active"],
            info["max_score"],
            info["legends_count"],
            winners_text,
        ]
        for col, val in enumerate(values, start=1):
            ws.cell(row=row_i, column=col).value = val
        ws.row_dimensions[row_i].height = max(28, 18 * max(1, info["legends_count"]))
        row_i += 1

    ws.column_dimensions["H"].width = 34
    style_table(ws, 2, max(2, ws.max_row), 1, 8, header_row=2)
    if ws.max_row >= 3:
        add_bar_chart(ws, "مجموع نقاط كل يوم", 4, 2, 3, ws.max_row, 1, f"A{ws.max_row + 4}", width=12, height=7, chart_type="col")
        add_bar_chart(ws, "متوسط نقاط المشاركين", 5, 2, 3, ws.max_row, 1, f"H{ws.max_row + 4}", width=12, height=7, chart_type="col")

    # -------------------- 4) تحليل المشاركين ✅ كما هو --------------------
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
    add_bar_chart(ws, "مجموع نقاط المشاركين", 2, 1, 2, ws.max_row, 1, f"A{ws.max_row + 4}", width=14, height=8, chart_type="bar")

    # -------------------- 5) تحليل الكباتن ✅ حسب المشاركين --------------------
    ws = wb.create_sheet("تحليل الكباتن")
    sheet_setup(ws, "تحليل الكباتن")

    card(ws, "A3:C5", "👑 أفضل مشارك في الكباتن", f"{best_captain_participant}\n{captain_by_participant.get(best_captain_participant, {}).get('points', 0)} نقطة إضافية", PURPLE)
    top_cap_player = captain_player_count.most_common(1)[0][0] if captain_player_count else "-"
    card(ws, "D3:F5", "🔥 أكثر لاعب تم اختياره كابتن", f"{top_cap_player}\n{captain_player_count.get(top_cap_player, 0)} مرات", BLUE)
    top_success = "-"
    if ranking:
        top_success = max(ranking, key=lambda n: (safe_avg(captain_by_participant[n]["success"], captain_by_participant[n]["count"]), captain_by_participant[n]["points"]))
    card(ws, "G3:I5", "🎯 أعلى نجاح كابتن", f"{top_success}\n{safe_pct(captain_by_participant[top_success]['success'], captain_by_participant[top_success]['count']) if top_success != '-' else '0%'}", GREEN)

    section_title(ws, 7, 1, "أداء المشاركين في اختيارات الكابتن", 8, HEADER)
    headers = ["المشارك", "مرات اختيار كابتن", "كباتن جابوا نقاط", "نقاط الكابتن الإضافية", "أفضل كابتن اختاره", "نسبة نجاح الكابتن"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=8, column=col).value = h

    row_i = 9
    for name in ranking:
        d = captain_by_participant[name]
        count = d["count"]
        best_caption = "-"
        if d["best_points"] > 0:
            best_caption = f"{d['best_player']} +{d['best_points']}"
        ws.cell(row=row_i, column=1).value = name
        ws.cell(row=row_i, column=2).value = count
        ws.cell(row=row_i, column=3).value = d["success"]
        ws.cell(row=row_i, column=4).value = d["points"]
        ws.cell(row=row_i, column=5).value = best_caption
        ws.cell(row=row_i, column=6).value = safe_pct(d["success"], count)
        row_i += 1

    style_table(ws, 8, row_i - 1, 1, 6, header_row=8)

    start2 = row_i + 3
    section_title(ws, start2, 1, "أكثر اللاعبين اختيارًا كابتن", 4, HEADER)
    for col, h in enumerate(["اللاعب", "مرات كابتن", "نقاط كابتن", "متوسط نقاط كابتن"], start=1):
        ws.cell(row=start2 + 1, column=col).value = h

    r = start2 + 2
    for player, count in captain_player_count.most_common():
        pts = captain_player_points[player]
        ws.cell(row=r, column=1).value = player
        ws.cell(row=r, column=2).value = count
        ws.cell(row=r, column=3).value = pts
        ws.cell(row=r, column=4).value = safe_avg(pts, count)
        r += 1

    style_table(ws, start2 + 1, max(start2 + 1, r - 1), 1, 4, header_row=start2 + 1)
    if row_i > 9:
        add_bar_chart(ws, "نقاط الكابتن الإضافية حسب المشارك", 4, 8, 9, row_i - 1, 1, f"A{row_i + 3}", width=13, height=8, chart_type="bar")
    if r > start2 + 2:
        add_bar_chart(ws, "مرات اختيار اللاعب كابتن", 2, start2 + 1, start2 + 2, min(r - 1, start2 + 16), 1, f"H{row_i + 3}", width=13, height=8, chart_type="col")

    # -------------------- 6) تحليل الحراس ✅ بدون تكرار الكلين شيت --------------------
    ws = wb.create_sheet("تحليل الحراس")
    sheet_setup(ws, "تحليل الحراس")

    top_keeper_points = best_keeper
    top_keeper_choice = stats["keeper_choice_count"].most_common(1)[0][0] if stats["keeper_choice_count"] else "-"
    best_keeper_rate = "-"
    if keeper_rounds:
        best_keeper_rate = max(keeper_rounds.keys(), key=lambda k: (len(keeper_clean_rounds[k]) / len(keeper_rounds[k]) if keeper_rounds[k] else 0, len(keeper_clean_rounds[k])))

    card(ws, "A3:C5", "🧤 أفضل حارس نقاط", f"{top_keeper_points}\n{stats['keeper_points_by_keeper'].get(top_keeper_points, 0)} نقطة للمشاركين", BLUE)
    card(ws, "D3:F5", "👥 أكثر حارس اختيارًا", f"{top_keeper_choice}\n{stats['keeper_choice_count'].get(top_keeper_choice, 0)} اختيار", PURPLE)
    card(ws, "G3:I5", "🎯 أعلى نسبة نجاح", f"{best_keeper_rate}\n{safe_pct(len(keeper_clean_rounds[best_keeper_rate]), len(keeper_rounds[best_keeper_rate])) if best_keeper_rate != '-' else '0%'}", GREEN)

    section_title(ws, 7, 1, "ملخص الحراس في البطولة", 8, HEADER)
    headers = ["الحارس", "إجمالي الاختيارات", "عدد المشاركين الذين اختاروه", "مرات الكلين شيت", "إجمالي نقاط المشاركين منه", "نسبة النجاح"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=8, column=col).value = h

    r = 9
    keeper_list = sorted(keeper_rounds.keys(), key=lambda k: (stats["keeper_points_by_keeper"][k], stats["keeper_choice_count"][k]), reverse=True)
    for keeper in keeper_list:
        rounds = sorted(keeper_rounds[keeper])
        clean_count = len(keeper_clean_rounds[keeper])
        ws.cell(row=r, column=1).value = keeper
        ws.cell(row=r, column=2).value = stats["keeper_choice_count"][keeper]
        ws.cell(row=r, column=3).value = len(keeper_participants[keeper])
        ws.cell(row=r, column=4).value = clean_count
        ws.cell(row=r, column=5).value = stats["keeper_points_by_keeper"][keeper]
        ws.cell(row=r, column=6).value = safe_pct(clean_count, len(rounds))
        r += 1

    style_table(ws, 8, max(8, r - 1), 1, 6, header_row=8)

    start2 = r + 3
    section_title(ws, start2, 1, "تفصيل الحراس حسب الجولة", 6, HEADER)
    headers = ["الجولة", "الحارس", "عدد من اختاروه", "كلين شيت؟", "نقاط الحارس الأساسية", "إجمالي نقاط المشاركين"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=start2 + 1, column=col).value = h

    rr = start2 + 2
    for (day, keeper), d in sorted(keeper_round_detail.items(), key=lambda x: (x[0][0], x[0][1])):
        ws.cell(row=rr, column=1).value = day
        ws.cell(row=rr, column=2).value = keeper
        ws.cell(row=rr, column=3).value = d["selected"]
        ws.cell(row=rr, column=4).value = "نعم" if d["clean"] else "لا"
        ws.cell(row=rr, column=5).value = 5 if d["clean"] else 0
        ws.cell(row=rr, column=6).value = d["total_points"]
        rr += 1

    style_table(ws, start2 + 1, max(start2 + 1, rr - 1), 1, 6, header_row=start2 + 1)
    if r > 9:
        add_bar_chart(ws, "إجمالي نقاط المشاركين من الحراس", 5, 8, 9, min(r - 1, 23), 1, f"A{rr + 3}", width=13, height=8, chart_type="bar")

    # -------------------- 7) تحليل اللاعبين ✅ تأثير اللاعبين --------------------
    ws = wb.create_sheet("تحليل اللاعبين")
    sheet_setup(ws, "تحليل اللاعبين")

    if top_round_player:
        (top_day, top_player), top_info = top_round_player
        card(ws, "A3:C6", "🔥 أكثر لاعب وزّع نقاط على المشاركين في جولة واحدة", f"{top_player} — اليوم {top_day}\nوزّع {top_info['distributed']} نقطة على المشاركين في هذه الجولة", RED)
    else:
        card(ws, "A3:C6", "🔥 أكثر لاعب وزّع نقاط على المشاركين في جولة واحدة", "-", RED)

    card(ws, "D3:F6", "⚽ أعلى لاعب نقاط فعلية في البطولة", f"{best_actual_player}\n{player_actual_total.get(best_actual_player, 0)} نقطة", BLUE)
    card(ws, "G3:I6", "👥 أكثر لاعب تم اختياره في البطولة", f"{most_selected_player}\n{player_choice_count.get(most_selected_player, 0)} اختيار", PURPLE)
    card(ws, "J3:L6", "👑 أكثر لاعب تم اختياره كابتن في البطولة", f"{most_captained_player}\n{player_captain_count.get(most_captained_player, 0)} مرة", GOLD)

    section_title(ws, 8, 1, "تفصيل تأثير اللاعبين على المشاركين في كل جولة", 6, HEADER)
    headers = ["الجولة", "اللاعب", "نقاط اللاعب في الجولة", "عدد اللي اختاروه", "عدد اللي كبتنوه", "النقاط التي وزعها على المشاركين في هذه الجولة"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=9, column=col).value = h

    r = 10
    round_rows = sorted(player_round_detail.items(), key=lambda x: (x[0][0], x[1]["distributed"]), reverse=False)
    round_rows = sorted(round_rows, key=lambda x: (x[0][0], -x[1]["distributed"], x[0][1]))
    for (day, player), d in round_rows:
        ws.cell(row=r, column=1).value = day
        ws.cell(row=r, column=2).value = player
        ws.cell(row=r, column=3).value = d["base_points"]
        ws.cell(row=r, column=4).value = d["selected"]
        ws.cell(row=r, column=5).value = d["captains"]
        ws.cell(row=r, column=6).value = d["distributed"]
        r += 1

    style_table(ws, 9, max(9, r - 1), 1, 6, header_row=9)

    start2 = r + 3
    section_title(ws, start2, 1, "ملخص اللاعبين في البطولة", 6, HEADER)
    headers = ["اللاعب", "إجمالي الاختيارات", "عدد المشاركين الذين اختاروه", "مرات كابتن", "نقاط اللاعب الفعلية في البطولة", "إجمالي النقاط التي أفاد بها المشاركين في البطولة"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=start2 + 1, column=col).value = h

    rr = start2 + 2
    player_list = sorted(player_choice_count.keys(), key=lambda p: (player_total_distributed[p], player_actual_total[p], player_choice_count[p]), reverse=True)
    for player in player_list:
        ws.cell(row=rr, column=1).value = player
        ws.cell(row=rr, column=2).value = player_choice_count[player]
        ws.cell(row=rr, column=3).value = len(player_participants[player])
        ws.cell(row=rr, column=4).value = player_captain_count[player]
        ws.cell(row=rr, column=5).value = player_actual_total[player]
        ws.cell(row=rr, column=6).value = player_total_distributed[player]
        rr += 1

    style_table(ws, start2 + 1, max(start2 + 1, rr - 1), 1, 6, header_row=start2 + 1)
    if r > 10:
        add_bar_chart(ws, "أعلى تأثير للاعب في جولة", 6, 9, 10, min(r - 1, 24), 2, f"A{rr + 4}", width=14, height=8, chart_type="bar")
    if rr > start2 + 2:
        add_bar_chart(ws, "إجمالي تأثير اللاعبين في البطولة", 6, start2 + 1, start2 + 2, min(rr - 1, start2 + 16), 1, f"H{rr + 4}", width=14, height=8, chart_type="bar")

    # -------------------- 8) تفصيل اختيارات اللاعبين ✅ حسب الجولة --------------------
    ws = wb.create_sheet("تفصيل اختيارات اللاعبين")
    sheet_setup(ws, "تفصيل اختيارات اللاعبين")

    headers = ["الجولة", "المشارك", "اللاعب", "كابتن؟", "نقاط اللاعب", "نقاط الكابتن", "الإجمالي", "مصدر النقاط"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=2, column=col).value = h

    round_fills = ["EAF1FF", "E7F8F1", "EFE7FF", "FFF6D6", "E8F7FF", "FFEAF4", "FFF0E0", "F3F4F6", "E6FFFA", "F5F0FF"]
    r = 3
    for row in sorted(selection_detail_rows, key=lambda x: (x[0], x[1], 0 if "نعم" in str(x[3]) else 1, x[2])):
        day = int(row[0]) if str(row[0]).isdigit() else 0
        fill = PatternFill("solid", fgColor=round_fills[(day - 1) % len(round_fills)] if day else "FFFFFF")
        for col, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=col)
            cell.value = val
            cell.fill = fill
        if "نعم" in str(row[3]):
            ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor=L_GOLD)
            ws.cell(row=r, column=4).font = Font(bold=True, color=TEXT)
        if row[6] and row[6] > 0:
            ws.cell(row=r, column=7).font = Font(bold=True, color=NAVY)
        r += 1

    style_table(ws, 2, max(2, r - 1), 1, 8, header_row=2)
    # نعيد تطبيق ألوان الجولات بعد تنسيق الجدول
    for rr2 in range(3, r):
        day = ws.cell(row=rr2, column=1).value
        try:
            day_int = int(day)
            fill = PatternFill("solid", fgColor=round_fills[(day_int - 1) % len(round_fills)])
            for c in range(1, 9):
                ws.cell(row=rr2, column=c).fill = fill
        except Exception:
            pass
        if "نعم" in str(ws.cell(row=rr2, column=4).value):
            ws.cell(row=rr2, column=4).fill = PatternFill("solid", fgColor=L_GOLD)
            ws.cell(row=rr2, column=4).font = Font(bold=True, color=TEXT)
        if score_to_int(ws.cell(row=rr2, column=7).value) > 0:
            ws.cell(row=rr2, column=7).font = Font(bold=True, color=NAVY)

    # -------------------- 9) سجل الأساطير ✅ --------------------
    ws = wb.create_sheet("سجل الأساطير")
    sheet_setup(ws, "سجل الأساطير")

    headers = ["الجولة", "أسطورة اليوم", "عدد الأساطير", "أعلى نقاط"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=2, column=col).value = h

    r = 3
    for day in days:
        info = day_summary[day]
        winners = info["winners"]
        ws.cell(row=r, column=1).value = f"اليوم {day}"
        ws.cell(row=r, column=2).value = " + ".join(winners) if winners else "-"
        ws.cell(row=r, column=3).value = info["legends_count"]
        ws.cell(row=r, column=4).value = info["max_score"]
        r += 1

    style_table(ws, 2, max(2, r - 1), 1, 4, header_row=2)
    ws.column_dimensions["B"].width = 42
    if r > 3:
        add_bar_chart(ws, "أعلى نقاط أسطورة اليوم", 4, 2, 3, r - 1, 1, "F3", width=13, height=7, chart_type="col")

    # تنسيق عام نهائي
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False
        for col in range(1, ws.max_column + 1):
            letter = get_column_letter(col)
            current_width = ws.column_dimensions[letter].width or 12
            ws.column_dimensions[letter].width = max(min(current_width, 42), 14)
        for row in range(1, ws.max_row + 1):
            if not ws.row_dimensions[row].height:
                ws.row_dimensions[row].height = 22

    # عروض خاصة
    if "تحليل اللاعبين" in wb.sheetnames:
        ws = wb["تحليل اللاعبين"]
        ws.column_dimensions["F"].width = 42
    if "تفصيل اختيارات اللاعبين" in wb.sheetnames:
        ws = wb["تفصيل اختيارات اللاعبين"]
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 24
        ws.column_dimensions["H"].width = 26
    if "تحليل الأيام" in wb.sheetnames:
        ws = wb["تحليل الأيام"]
        ws.column_dimensions["H"].width = 36

    wb.save(file_name)
    return file_name, stats


# ============================================================
# أوامر إضافية لفانتزي المصيف
# أضف هذا القسم قبل: async def dashboard
# ============================================================

PENDING_IMPORTS = {}
INVALID_SELECTIONS_GLOBAL = {"", "لم يشارك", "تم حجب المشاركة / تأخير"}

# استبدل دالة is_no_participation القديمة بهذا التعريف
# أو اترك هذا التعريف بعد القديمة عشان يغطيها.
def is_no_participation(value):
    value = normalize_name(value)
    return value in INVALID_SELECTIONS_GLOBAL


def load_workbook_safely(path, data_only=True):
    """يفتح ملفات الإكسل حتى لو فيها مشكلة styles مثل family val > 14."""
    import zipfile
    import tempfile
    import re as _re

    try:
        return load_workbook(path, data_only=data_only)
    except Exception:
        fixed_path = os.path.join(tempfile.gettempdir(), f"fixed_{os.path.basename(path)}")
        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(fixed_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/styles.xml":
                    text = data.decode("utf-8", errors="ignore")
                    text = _re.sub(r'<family val="(?:1[5-9]|[2-9][0-9]+)"\s*/>', '<family val="2"/>', text)
                    data = text.encode("utf-8")
                zout.writestr(item, data)
        return load_workbook(fixed_path, data_only=data_only)


def cell_text(value):
    return normalize_name(value)


def cell_int(value, default=0):
    try:
        if value is None or value == "":
            return default
        return int(float(value))
    except Exception:
        return default


def parse_import_excel(path):
    """
    يستورد ملف الإكسل كامل مرة وحدة.
    يدعم صفحات باسم: يوم 1، يوم 2، ...
    ويسحب:
    - جدول النتائج من الأعمدة A:B:C
    - جدول المشاركين من عند عمود "المشارك"
    ويتجاهل الأيام الفاضية بالكامل حتى لا ينشئ أيام وهمية.
    """
    wb = load_workbook_safely(path, data_only=True)
    imported = {}

    TRUE_VALUES = {"نعم", "yes", "Yes", "YES", "صح", "✓", "✅", "1", "true", "True"}

    for ws in wb.worksheets:
        m = re.search(r"يوم\s*(\d+)", str(ws.title))
        if not m:
            continue

        day = int(m.group(1))

        # نبحث عن خانة "المشارك" في أول الصفوف
        header_row = None
        participant_col = None
        for r in range(1, min(ws.max_row, 25) + 1):
            for c in range(1, min(ws.max_column, 25) + 1):
                if cell_text(ws.cell(r, c).value) == "المشارك":
                    header_row = r
                    participant_col = c
                    break
            if header_row:
                break

        if not header_row or not participant_col:
            continue

        lineups = {}
        active_found = False

        # المشاركون: المشارك، الحارس، اللاعب 1، اللاعب 2، اللاعب 3، الكابتن
        for r in range(header_row + 1, ws.max_row + 1):
            participant = cell_text(ws.cell(r, participant_col).value)
            if not participant or participant not in PARTICIPANTS:
                continue

            values = [cell_text(ws.cell(r, participant_col + offset).value) for offset in range(1, 6)]

            # لو الصف فاضي أو كله عدم مشاركة
            if all(is_no_participation(v) for v in values):
                values = ["لم يشارك"] * 5
            else:
                active_found = True

            lineups[participant] = values

        # النتائج: غالبًا في A:B:C = اسم اللاعب | عدد الأهداف | كلين شيت؟
        goals_count = defaultdict(int)
        clean_sheets = []

        for r in range(header_row + 1, ws.max_row + 1):
            name = cell_text(ws.cell(r, 1).value)
            goals_raw = ws.cell(r, 2).value
            clean_raw = cell_text(ws.cell(r, 3).value)

            if not name or is_no_participation(name):
                continue

            goals = cell_int(goals_raw, 0)
            if goals > 0:
                goals_count[name] += goals

            if clean_raw in TRUE_VALUES:
                if name not in clean_sheets:
                    clean_sheets.append(name)

        # لا نستورد الأيام الفاضية بالكامل، مثل يوم 7 وما بعده إذا ما انلعبت
        if not active_found and not goals_count and not clean_sheets:
            continue

        if lineups:
            imported[day] = {
                "lineups": lineups,
                "goals_count": dict(goals_count),
                "clean_sheets": clean_sheets,
            }

    return imported

def import_summary_text(imported):
    lines = ["تم قراءة الملف ✅", "", "الأيام الموجودة:"]
    if not imported:
        lines.append("ما لقيت أيام قابلة للاستيراد.")
        return "\n".join(lines)

    for day in sorted(imported):
        data = imported[day]
        participants = len(data["lineups"])
        has_results = bool(data["goals_count"] or data["clean_sheets"])
        status = "نتائج موجودة" if has_results else "بدون نتائج"
        exists = " — موجود مسبقًا وسيتم استبداله" if os.path.exists(excel_file(day)) else ""
        lines.append(f"اليوم {day}: {participants} مشارك + {status}{exists}")

    lines.append("")
    lines.append("للاعتماد اكتب:")
    lines.append("/اعتماد_استيراد")
    lines.append("")
    lines.append("للإلغاء اكتب:")
    lines.append("/إلغاء_استيراد")
    return "\n".join(lines)


async def import_excel_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    if not document:
        await update.message.reply_text(
            "أرسل ملف الإكسل كملف وليس صورة، واكتب معه في التعليق:\n/استيراد_ملف"
        )
        return

    filename = document.file_name or "import.xlsx"
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        await update.message.reply_text("الملف لازم يكون Excel بصيغة .xlsx أو .xlsm")
        return

    os.makedirs("imports", exist_ok=True)
    local_path = os.path.join("imports", f"import_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{filename}")

    tg_file = await context.bot.get_file(document.file_id)
    await tg_file.download_to_drive(local_path)

    try:
        imported = parse_import_excel(local_path)
    except Exception as e:
        await update.message.reply_text(f"صار خطأ أثناء قراءة الإكسل ❌\n{e}")
        return

    if not imported:
        await update.message.reply_text("ما قدرت أستخرج أيام من الملف. تأكد أن الصفحات باسم: يوم 1، يوم 2 ...")
        return

    chat_id = update.effective_chat.id
    PENDING_IMPORTS[chat_id] = {"path": local_path, "data": imported}
    await update.message.reply_text(import_summary_text(imported))


async def approve_import(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    pending = PENDING_IMPORTS.get(chat_id)
    if not pending:
        await update.message.reply_text("ما فيه استيراد معلق. أرسل ملف الإكسل مع /استيراد_ملف أولًا.")
        return

    imported = pending["data"]
    files_to_backup = [excel_file(day) for day in imported.keys() if os.path.exists(excel_file(day))]
    if os.path.exists(LOCKED_FILE):
        files_to_backup.append(LOCKED_FILE)
    if files_to_backup:
        backup_files("before_import_excel", files=files_to_backup)

    locked = load_locked_days()
    saved_days = []
    warnings = []

    for day in sorted(imported):
        data = imported[day]
        file_name, missing = update_day_data(str(day), data["lineups"])
        if missing:
            warnings.extend([f"اليوم {day}: مشارك غير موجود بالقائمة: {m}" for m in missing])
        calculate_points(str(day), data["goals_count"], data["clean_sheets"])
        locked.add(str(day))
        saved_days.append(day)

    save_locked_days(locked)
    PENDING_IMPORTS.pop(chat_id, None)

    msg = [
        "تم اعتماد الاستيراد ✅",
        f"الأيام المحفوظة: {', '.join(map(str, saved_days))}",
        "تم حساب النتائج وقفل الأيام المستوردة.",
    ]
    if warnings:
        msg.append("\nتنبيهات:")
        msg.extend(warnings[:20])
    await update.message.reply_text("\n".join(msg))


async def cancel_import(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    if chat_id in PENDING_IMPORTS:
        PENDING_IMPORTS.pop(chat_id, None)
        await update.message.reply_text("تم إلغاء الاستيراد ✅")
    else:
        await update.message.reply_text("ما فيه استيراد معلق.")


async def backup_zip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import zipfile

    files = [f for f in os.listdir(".") if re.match(r"fantasy_day_\d+\.xlsx$", f)]
    if os.path.exists(LOCKED_FILE):
        files.append(LOCKED_FILE)

    if not files:
        await update.message.reply_text("ما فيه ملفات أيام عشان أسوي نسخة احتياطية.")
        return

    zip_name = f"fantasy_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(files):
            z.write(f)

    with open(zip_name, "rb") as f:
        await update.message.reply_document(document=f, filename=zip_name, caption="نسخة احتياطية لأيام الفانتزي ✅")


async def restore_backup_zip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import zipfile

    document = update.message.document
    if not document:
        await update.message.reply_text("أرسل ملف ZIP واكتب معه في التعليق:\n/استرجاع_نسخة")
        return

    filename = document.file_name or "backup.zip"
    if not filename.lower().endswith(".zip"):
        await update.message.reply_text("الملف لازم يكون ZIP.")
        return

    os.makedirs("restore_uploads", exist_ok=True)
    local_path = os.path.join("restore_uploads", f"restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{filename}")
    tg_file = await context.bot.get_file(document.file_id)
    await tg_file.download_to_drive(local_path)

    backup_files("before_restore_zip")
    restored = []
    try:
        with zipfile.ZipFile(local_path, "r") as z:
            for name in z.namelist():
                base = os.path.basename(name)
                if re.match(r"fantasy_day_\d+\.xlsx$", base) or base == LOCKED_FILE:
                    with z.open(name) as src, open(base, "wb") as dst:
                        dst.write(src.read())
                    restored.append(base)
    except Exception as e:
        await update.message.reply_text(f"صار خطأ في استرجاع النسخة ❌\n{e}")
        return

    if not restored:
        await update.message.reply_text("ملف ZIP ما فيه ملفات أيام صالحة.")
        return

    await update.message.reply_text("تم استرجاع النسخة ✅\n" + "\n".join(sorted(restored)))


async def clean_days(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """ينظف القفل من أيام ما لها ملف فعلي."""
    locked = load_locked_days()
    existing = {str(d) for d in get_existing_days(1, 99)}
    cleaned = {d for d in locked if d in existing}
    removed = sorted(locked - cleaned, key=lambda x: int(x))
    save_locked_days(cleaned)
    if removed:
        await update.message.reply_text("تم تنظيف الأيام الوهمية ✅\nأزيلت من القفل: " + "، ".join(removed))
    else:
        await update.message.reply_text("ما فيه أيام وهمية. كل شيء تمام ✅")



# -------------------- أوامر تيليجرام --------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "الأوامر الأساسية:\n"
        "/اضافه 5\n"
        "/نتائج 5\n"
        "/احصائيات\n"
        "/احصائيات 1 6\n"
        "/ترتيب_نص\n\n"
        "أوامر الفحص:\n"
        "/الأيام\n"
        "/فحص 5\n"
        "/مشاركين 5\n"
        "/اسطورة 5\n"
        "/مقارنة 4 5\n\n"
        "أوامر الاستيراد والنسخ:\n"
        "/استيراد_ملف  — أرسلها كتعليق مع ملف الإكسل\n"
        "/اعتماد_استيراد\n"
        "/إلغاء_استيراد\n"
        "/نسخة_احتياطية\n"
        "/استرجاع_نسخة  — أرسلها كتعليق مع ملف ZIP\n"
        "/تنظيف_الأيام\n\n"
        "أوامر الأمان:\n"
        "/مسح_نتائج 5\n"
        "/مسح_يوم 5\n"
        "/مسح_الكل تأكيد\n"
        "/استرجاع_آخر\n"
        "/قفل_يوم 5\n"
        "/فتح_يوم 5"
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
        if start_day > end_day:
            start_day, end_day = end_day, start_day
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
        if start_day > end_day:
            start_day, end_day = end_day, start_day
    elif len(nums) == 1:
        start_day, end_day = 1, nums[0]
    else:
        start_day, end_day = 1, 31
    stats = collect_stats(start_day, end_day)
    await update.message.reply_text(build_ranking_text(stats, start_day, end_day))


async def list_days(update: Update, context: ContextTypes.DEFAULT_TYPE):
    days = get_existing_days(1, 99)
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
        nums = get_numbers(update.message.text or update.message.caption or "")

        if len(nums) >= 2:
            start_day, end_day = nums[0], nums[1]
            if start_day > end_day:
                start_day, end_day = end_day, start_day
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
            "تطور النقاط\n"
            "تحليل الأيام\n"
            "تحليل المشاركين\n"
            "تحليل الكباتن\n"
            "تحليل الحراس\n"
            "تحليل اللاعبين\n"
            "تفصيل اختيارات اللاعبين\n"
            "سجل الأساطير"
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

    # استيراد ملف Excel كامل مرة واحدة
    app.add_handler(MessageHandler(filters.Document.ALL & filters.CaptionRegex(r"^/استيراد_ملف"), import_excel_file))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استيراد_ملف"), import_excel_file))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعتماد_استيراد"), approve_import))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إلغاء_استيراد"), cancel_import))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الغاء_استيراد"), cancel_import))

    # النسخ الاحتياطي والاسترجاع
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نسخة_احتياطية"), backup_zip))
    app.add_handler(MessageHandler(filters.Document.ALL & filters.CaptionRegex(r"^/استرجاع_نسخة"), restore_backup_zip))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استرجاع_نسخة"), restore_backup_zip))

    # تنظيف الأيام الوهمية من القفل
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تنظيف_الأيام"), clean_days))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تنظيف_الايام"), clean_days))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اضافه"), add_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج"), results_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الترتيب_العام"), overall))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/ترتيب_نص"), ranking_text))

    # الإحصائيات — يدعم:
    # /احصائيات
    # /احصائيات 1 6
    # /احصائيات 6 1
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
