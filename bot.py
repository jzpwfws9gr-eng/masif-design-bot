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
CUP_FILE = "cup_state.json"

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
        ) or filename in ("overall_ranking.xlsx", "fantasy_dashboard.xlsx", LOCKED_FILE, CUP_FILE):
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
LAST_UPLOADED_FILES = {}
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


async def _download_document_to_folder(update: Update, context: ContextTypes.DEFAULT_TYPE, folder="uploads"):
    """يحفظ أي ملف مرسل ويرجع المسار المحلي."""
    document = update.message.document
    if not document:
        return None, None

    filename = document.file_name or "uploaded_file"
    os.makedirs(folder, exist_ok=True)
    safe_name = re.sub(r"[^A-Za-z0-9_.\-\u0600-\u06FF]+", "_", filename)
    local_path = os.path.join(folder, f"{update.effective_chat.id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_name}")

    tg_file = await context.bot.get_file(document.file_id)
    await tg_file.download_to_drive(local_path)
    return local_path, filename


async def remember_last_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """يحفظ آخر ملف Excel أو ZIP أرسله المستخدم، عشان نقدر نستخدمه بعدين بدون تعليق."""
    local_path, filename = await _download_document_to_folder(update, context, "uploads")
    if not local_path:
        return

    chat_id = update.effective_chat.id
    LAST_UPLOADED_FILES.setdefault(chat_id, {})
    lower = (filename or "").lower()
    caption = (update.message.caption or "").strip()

    if lower.endswith((".xlsx", ".xlsm")):
        LAST_UPLOADED_FILES[chat_id]["excel"] = local_path
        if caption.startswith("/استيراد_ملف"):
            await _run_import_from_path(update, context, local_path)
        else:
            await update.message.reply_text(
                "وصل ملف الإكسل ✅\n"
                "اكتب الآن:\n"
                "/استيراد_ملف"
            )
        return

    if lower.endswith(".zip"):
        LAST_UPLOADED_FILES[chat_id]["zip"] = local_path
        if caption.startswith("/استرجاع_نسخة"):
            await _run_restore_from_zip_path(update, context, local_path)
        else:
            await update.message.reply_text(
                "وصل ملف ZIP ✅\n"
                "للاسترجاع اكتب الآن:\n"
                "/استرجاع_نسخة"
            )
        return

    await update.message.reply_text("وصل الملف، لكن أحتاج Excel أو ZIP فقط.")


async def _run_import_from_path(update: Update, context: ContextTypes.DEFAULT_TYPE, local_path):
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


async def import_excel_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    يدعم طريقتين:
    1) إرسال ملف الإكسل بتعليق /استيراد_ملف
    2) إرسال ملف الإكسل لحاله، ثم كتابة /استيراد_ملف
    """
    document = update.message.document
    chat_id = update.effective_chat.id

    if document:
        filename = document.file_name or "import.xlsx"
        if not filename.lower().endswith((".xlsx", ".xlsm")):
            await update.message.reply_text("الملف لازم يكون Excel بصيغة .xlsx أو .xlsm")
            return
        local_path, _ = await _download_document_to_folder(update, context, "imports")
        LAST_UPLOADED_FILES.setdefault(chat_id, {})["excel"] = local_path
        await _run_import_from_path(update, context, local_path)
        return

    local_path = LAST_UPLOADED_FILES.get(chat_id, {}).get("excel")
    if not local_path or not os.path.exists(local_path):
        await update.message.reply_text(
            "ما لقيت ملف Excel محفوظ.\n"
            "أرسل ملف الإكسل لحاله أولًا، وبعدها اكتب:\n"
            "/استيراد_ملف"
        )
        return

    await _run_import_from_path(update, context, local_path)


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


async def _run_restore_from_zip_path(update: Update, context: ContextTypes.DEFAULT_TYPE, local_path):
    import zipfile

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


async def restore_backup_zip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """يدعم إرسال ZIP لحاله ثم /استرجاع_نسخة، أو ZIP بتعليق."""
    document = update.message.document
    chat_id = update.effective_chat.id

    if document:
        filename = document.file_name or "backup.zip"
        if not filename.lower().endswith(".zip"):
            await update.message.reply_text("الملف لازم يكون ZIP.")
            return
        local_path, _ = await _download_document_to_folder(update, context, "restore_uploads")
        LAST_UPLOADED_FILES.setdefault(chat_id, {})["zip"] = local_path
        await _run_restore_from_zip_path(update, context, local_path)
        return

    local_path = LAST_UPLOADED_FILES.get(chat_id, {}).get("zip")
    if not local_path or not os.path.exists(local_path):
        await update.message.reply_text(
            "ما لقيت ملف ZIP محفوظ.\n"
            "أرسل ملف ZIP لحاله أولًا، وبعدها اكتب:\n"
            "/استرجاع_نسخة"
        )
        return

    await _run_restore_from_zip_path(update, context, local_path)


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
        "/استيراد_ملف — أرسل ملف الإكسل لحاله ثم اكتب الأمر\n"
        "/اعتماد_استيراد\n"
        "/إلغاء_استيراد\n"
        "/نسخة_احتياطية\n"
        "/استرجاع_نسخة — أرسل ملف ZIP لحاله ثم اكتب الأمر\n"
        "/تنظيف_الأيام\n\n"
        "أوامر الأمان:\n"
        "/مسح_نتائج 5\n"
        "/مسح_يوم 5\n"
        "/مسح_الكل تأكيد\n"
        "/استرجاع_آخر\n"
        "/قفل_يوم 5\n"
        "/فتح_يوم 5"
    )




# ============================================================
# V19 — كأس المصيف اليدوي + تصميم مواجهات الكأس
# ============================================================

CUP_ROUNDS = [
    ("r12", "دور 12"),
    ("qf", "ربع النهائي"),
    ("sf", "نصف النهائي"),
    ("final", "النهائي"),
]

def load_cup_state():
    if not os.path.exists(CUP_FILE):
        return {"active": False}
    try:
        with open(CUP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {"active": False}
        return data
    except Exception:
        return {"active": False}

def save_cup_state(state):
    with open(CUP_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def cup_round_key_by_offset(offset):
    if 0 <= offset < len(CUP_ROUNDS):
        return CUP_ROUNDS[offset][0]
    return None

def cup_round_name(round_key):
    for key, name in CUP_ROUNDS:
        if key == round_key:
            return name
    return str(round_key or "")

def cup_next_round(round_key):
    keys = [x[0] for x in CUP_ROUNDS]
    if round_key not in keys:
        return None
    idx = keys.index(round_key)
    return keys[idx + 1] if idx + 1 < len(keys) else None

def cup_previous_ranking_before(start_day):
    try:
        prev_end = max(0, int(start_day) - 1)
    except Exception:
        prev_end = 0

    if prev_end >= 1:
        stats = collect_stats(1, prev_end)
        ranking = [n for n in stats.get("ranking", []) if n in PARTICIPANTS]
        missing = [n for n in PARTICIPANTS if n not in ranking]
        return ranking + missing
    return PARTICIPANTS[:]

def cup_seed_number(state, participant):
    try:
        return int(state.get("seed_by", {}).get(participant, 99))
    except Exception:
        return 99

def cup_make_match(a, b, match_id=None):
    return {
        "id": match_id or "",
        "a": a,
        "b": b,
        "score_a": None,
        "score_b": None,
        "active_a": None,
        "active_b": None,
        "winner": None,
        "note": "",
    }

def cup_initial_state(start_day):
    seeds = cup_previous_ranking_before(start_day)[:12]
    while len(seeds) < 12:
        seeds.append(f"مشارك {len(seeds)+1}")

    seed_by = {name: i + 1 for i, name in enumerate(seeds)}

    # دور 12: أول 4 عندهم راحة، والباقي يلعبون
    r12_matches = [
        cup_make_match(seeds[4], seeds[11], "R12-1"),  # 5 vs 12
        cup_make_match(seeds[7], seeds[8], "R12-2"),   # 8 vs 9
        cup_make_match(seeds[5], seeds[10], "R12-3"),  # 6 vs 11
        cup_make_match(seeds[6], seeds[9], "R12-4"),   # 7 vs 10
    ]

    return {
        "active": True,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "start_day": int(start_day),
        "end_day": int(start_day) + 3,
        "seeds": seeds,
        "seed_by": seed_by,
        "champion": None,
        "rounds": {
            "r12": {"day": int(start_day), "name": "دور 12", "matches": r12_matches, "completed": False},
            "qf": {"day": int(start_day) + 1, "name": "ربع النهائي", "matches": [], "completed": False},
            "sf": {"day": int(start_day) + 2, "name": "نصف النهائي", "matches": [], "completed": False},
            "final": {"day": int(start_day) + 3, "name": "النهائي", "matches": [], "completed": False},
        },
    }

def cup_scores_for_day(day):
    rows = read_day_rows(day)
    scores = {}
    active = {}
    for row in rows:
        name = row.get("participant")
        if not name:
            continue
        scores[name] = int(row.get("total", 0) or 0)
        active[name] = bool(row.get("participated"))
    for name in PARTICIPANTS:
        scores.setdefault(name, 0)
        active.setdefault(name, False)
    return scores, active

def cup_decide_winner(state, match, scores, active):
    a = match.get("a", "")
    b = match.get("b", "")
    sa = int(scores.get(a, 0) or 0)
    sb = int(scores.get(b, 0) or 0)
    aa = bool(active.get(a, False))
    ab = bool(active.get(b, False))

    # القواعد المعتمدة:
    # إذا واحد شارك والثاني لا: المشارك يتأهل حتى لو صفر
    # إذا الاثنين شاركوا: الأعلى نقاط يتأهل
    # إذا الاثنين لم يشاركوا: الأعلى بالتصنيف/البذر يتأهل
    if aa and not ab:
        winner = a
        note = "تأهل لأنه شارك"
    elif ab and not aa:
        winner = b
        note = "تأهل لأنه شارك"
    elif aa and ab:
        if sa > sb:
            winner = a
            note = "فاز بالنقاط"
        elif sb > sa:
            winner = b
            note = "فاز بالنقاط"
        else:
            # عند التعادل: الأعلى بذرًا يتأهل
            seed_a = cup_seed_number(state, a)
            seed_b = cup_seed_number(state, b)
            winner = a if seed_a <= seed_b else b
            note = "تعادل — تأهل الأعلى تصنيفًا"
    else:
        seed_a = cup_seed_number(state, a)
        seed_b = cup_seed_number(state, b)
        winner = a if seed_a <= seed_b else b
        note = "لم يشارك الطرفان — تأهل الأعلى تصنيفًا"

    match["score_a"] = sa
    match["score_b"] = sb
    match["active_a"] = aa
    match["active_b"] = ab
    match["winner"] = winner
    match["note"] = note
    return winner

def cup_build_next_round_matches(state, completed_round_key):
    seeds = state.get("seeds", [])
    rounds = state.get("rounds", {})

    if completed_round_key == "r12":
        r12 = rounds.get("r12", {}).get("matches", [])
        if len(r12) < 4 or len(seeds) < 12:
            return []
        # ترتيب ربع النهائي
        # Seed 1 vs Winner 8/9
        # Seed 4 vs Winner 5/12
        # Seed 2 vs Winner 7/10
        # Seed 3 vs Winner 6/11
        return [
            cup_make_match(seeds[0], r12[1].get("winner"), "QF-1"),
            cup_make_match(seeds[3], r12[0].get("winner"), "QF-2"),
            cup_make_match(seeds[1], r12[3].get("winner"), "QF-3"),
            cup_make_match(seeds[2], r12[2].get("winner"), "QF-4"),
        ]

    if completed_round_key == "qf":
        qf = rounds.get("qf", {}).get("matches", [])
        if len(qf) < 4:
            return []
        return [
            cup_make_match(qf[0].get("winner"), qf[1].get("winner"), "SF-1"),
            cup_make_match(qf[2].get("winner"), qf[3].get("winner"), "SF-2"),
        ]

    if completed_round_key == "sf":
        sf = rounds.get("sf", {}).get("matches", [])
        if len(sf) < 2:
            return []
        return [
            cup_make_match(sf[0].get("winner"), sf[1].get("winner"), "FINAL"),
        ]

    return []

def cup_process_day(day):
    state = load_cup_state()
    if not state.get("active"):
        return state, None, "لا يوجد كأس مفعّلة."

    try:
        day_i = int(day)
        start_day = int(state.get("start_day"))
    except Exception:
        return state, None, "رقم اليوم غير صحيح."

    if day_i < start_day or day_i > int(state.get("end_day", start_day + 3)):
        return state, None, "اليوم خارج فترة الكأس الحالية."

    round_key = cup_round_key_by_offset(day_i - start_day)
    if not round_key:
        return state, None, "اليوم خارج مراحل الكأس."

    round_data = state.get("rounds", {}).get(round_key, {})
    matches = round_data.get("matches", [])
    if not matches:
        return state, None, f"لا توجد مواجهات في {cup_round_name(round_key)}."

    if round_data.get("completed"):
        return state, round_key, "هذه المرحلة محسوبة مسبقًا."

    scores, active = cup_scores_for_day(day_i)
    for match in matches:
        cup_decide_winner(state, match, scores, active)

    round_data["completed"] = True
    state["rounds"][round_key] = round_data

    next_key = cup_next_round(round_key)
    if next_key:
        next_matches = cup_build_next_round_matches(state, round_key)
        state["rounds"][next_key]["matches"] = next_matches
    else:
        # النهائي انتهى
        winner = matches[0].get("winner") if matches else None
        state["champion"] = winner
        state["active"] = False

    save_cup_state(state)
    return state, round_key, "تم حساب مرحلة الكأس ✅"

def cup_pending_round_key(state):
    if not state.get("rounds"):
        return None
    for key, _name in CUP_ROUNDS:
        rd = state["rounds"].get(key, {})
        if rd.get("matches") and not rd.get("completed"):
            return key
    if state.get("champion"):
        return "champion"
    return None

def cup_results_text(state):
    if not state or not state.get("rounds"):
        return "لا توجد كأس محفوظة."
    lines = [
        "🏆 كأس المصيف",
        f"الفترة: اليوم {state.get('start_day')} إلى اليوم {state.get('end_day')}",
        "",
    ]

    for key, name in CUP_ROUNDS:
        rd = state.get("rounds", {}).get(key, {})
        matches = rd.get("matches", [])
        if not matches:
            continue
        lines.append(f"📌 {name} — اليوم {rd.get('day')}")
        for m in matches:
            a, b = m.get("a", ""), m.get("b", "")
            if m.get("winner"):
                lines.append(f"- {a} {m.get('score_a', 0)} × {m.get('score_b', 0)} {b} | المتأهل: {m.get('winner')}")
            else:
                lines.append(f"- {a} × {b}")
        lines.append("")

    if state.get("champion"):
        lines.append(f"🏆 بطل كأس المصيف: {state.get('champion')}")
    else:
        pending = cup_pending_round_key(state)
        if pending and pending != "champion":
            lines.append(f"القادم: {cup_round_name(pending)}")
    return "\n".join(lines).strip()

def create_cup_matches_image(state, round_key=None, title_suffix=None):
    ensure_generated_dir()
    if round_key is None:
        round_key = cup_pending_round_key(state)
    if not round_key:
        return None

    width = 1200
    if round_key == "champion":
        height = 850
        img, draw = design_canvas(None, width, height, "gold")
        draw_design_header(draw, width, "كأس المصيف 2026", "البطل", img)
        draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=112, accent="#F59E0B")
        champion = state.get("champion", "غير محدد")
        draw_text(draw, (width//2, 430), "🏆", get_font(82), fill="#FDE68A")
        draw_text(draw, (width//2, 535), champion, get_font(58), fill="#FFFFFF", max_width=850)
        draw_text(draw, (width//2, 625), "بطل كأس المصيف", get_font(36), fill="#FDE68A")
        footer_event(draw, width, height)
        path = os.path.join(GENERATED_DIR, "cup_champion.png")
        img.save(path, quality=95)
        return path

    rd = state.get("rounds", {}).get(round_key, {})
    matches = rd.get("matches", [])
    count = max(len(matches), 1)
    row_h, gap, name_size = v16_fit_row_metrics(count, "match")
    if count <= 2:
        row_h += 10
        name_size += 2

    content_h = count * row_h + max(0, count - 1) * gap
    height = max(780, 245 + content_h + 200)
    img, draw = design_canvas(None, width, height, "gold")
    subtitle = title_suffix or f"{cup_round_name(round_key)} — اليوم {rd.get('day', '')}"
    draw_design_header(draw, width, "كأس المصيف 2026", subtitle, img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=112, accent="#F59E0B")

    available_h = (fy2 - fy1) - 70
    y = fy1 + 38 + max(0, (available_h - content_h) // 2)
    for i, m in enumerate(matches, start=1):
        a = m.get("a") or "—"
        b = m.get("b") or "—"
        accent = "#F59E0B" if m.get("winner") else v16_accent(i)
        rounded_rect(draw, (92, y, width-92, y+row_h), radius=28, fill="#0B1020", outline=accent, width=2)
        cy = y + row_h//2

        # بذر/تصنيف اللاعب
        seed_a = cup_seed_number(state, a)
        seed_b = cup_seed_number(state, b)
        rounded_rect(draw, (width-190, cy-34, width-132, cy+34), radius=16, fill="#05070D", outline="#FDE68A80", width=1)
        rounded_rect(draw, (132, cy-34, 190, cy+34), radius=16, fill="#05070D", outline="#FDE68A80", width=1)
        draw_text(draw, (width-161, cy), str(seed_a if seed_a != 99 else "-"), get_font(28), fill="#FDE68A")
        draw_text(draw, (161, cy), str(seed_b if seed_b != 99 else "-"), get_font(28), fill="#FDE68A")

        a_fill = "#FDE68A" if m.get("winner") == a else "#FFFFFF"
        b_fill = "#FDE68A" if m.get("winner") == b else "#FFFFFF"
        draw_text(draw, (width-390, cy), a, get_font(name_size), fill=a_fill, max_width=330)
        draw_text(draw, (390, cy), b, get_font(name_size), fill=b_fill, max_width=330)

        if m.get("winner"):
            score = f"{m.get('score_a', 0)} - {m.get('score_b', 0)}"
            rounded_rect(draw, (width//2-110, cy-34, width//2+110, cy+34), radius=20, fill="#05070D", outline="#FDE68A", width=2)
            draw_text(draw, (width//2, cy), score, get_font(max(30, name_size+4)), fill="#FDE68A")
        else:
            draw_text(draw, (width//2, cy), "×", get_font(max(42, name_size+18)), fill="#FDE68A")

        y += row_h + gap

    footer_event(draw, width, height)
    path = os.path.join(GENERATED_DIR, f"cup_{round_key}.png")
    img.save(path, quality=95)
    return path

def cup_started_caption(state):
    seeds = state.get("seeds", [])
    bye = "، ".join(seeds[:4]) if len(seeds) >= 4 else "لا يوجد"
    return (
        f"تم بدء كأس المصيف 🏆\n"
        f"الفترة: اليوم {state.get('start_day')} إلى اليوم {state.get('end_day')}\n\n"
        f"اليوم {state.get('start_day')}: دور 12\n"
        f"اليوم {int(state.get('start_day'))+1}: ربع النهائي\n"
        f"اليوم {int(state.get('start_day'))+2}: نصف النهائي\n"
        f"اليوم {int(state.get('start_day'))+3}: النهائي\n\n"
        f"راحة دور 12: {bye}"
    )

async def start_cup_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nums = get_numbers(update.message.text)
    if not nums:
        await update.message.reply_text("اكتبها كذا:\n/بدء_الكاس 5")
        return
    start_day = int(nums[0])
    if start_day < 1:
        await update.message.reply_text("رقم اليوم غير صحيح.")
        return

    old = load_cup_state()
    if old.get("active"):
        await update.message.reply_text(
            "فيه كأس مفعّلة حاليًا.\n"
            "إذا تبي تلغيها اكتب: /الغاء_الكاس"
        )
        return

    backup_files(f"before_start_cup_{start_day}", files=[CUP_FILE] if os.path.exists(CUP_FILE) else [])
    state = cup_initial_state(start_day)
    save_cup_state(state)

    path = create_cup_matches_image(state, "r12", "دور 12")
    if path:
        await send_photo_path(update, path, cup_started_caption(state))
    else:
        await update.message.reply_text(cup_started_caption(state))

async def cup_status_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    state = load_cup_state()
    if not state.get("rounds"):
        await update.message.reply_text("لا توجد كأس محفوظة حاليًا.\nلبدء كأس جديد: /بدء_الكاس 5")
        return
    pending = cup_pending_round_key(state)
    status = "مفعّلة ✅" if state.get("active") else "منتهية ✅"
    msg = (
        f"🏆 حالة كأس المصيف\n"
        f"الحالة: {status}\n"
        f"الفترة: اليوم {state.get('start_day')} إلى اليوم {state.get('end_day')}\n"
    )
    if state.get("champion"):
        msg += f"البطل: {state.get('champion')}"
    elif pending:
        msg += f"المرحلة الحالية: {cup_round_name(pending)}"
    await update.message.reply_text(msg)

async def cup_results_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    state = load_cup_state()
    await update.message.reply_text(cup_results_text(state))

async def cancel_cup_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "تأكيد" not in (update.message.text or ""):
        await update.message.reply_text(
            "لإلغاء كأس المصيف الحالية اكتب:\n"
            "/الغاء_الكاس تأكيد\n\n"
            "تنبيه: هذا يلغي البطولة الحالية فقط، ولا يغير نقاط الفانتزي."
        )
        return

    state = load_cup_state()
    if not state.get("rounds"):
        await update.message.reply_text("ما فيه كأس محفوظة أصلًا.")
        return

    backup_files("before_cancel_cup", files=[CUP_FILE] if os.path.exists(CUP_FILE) else [])
    state["active"] = False
    state["cancelled_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_cup_state(state)
    await update.message.reply_text("تم إلغاء كأس المصيف الحالية ✅\nتقدر تبدأ كأس جديد بأمر:\n/بدء_الكاس 7")

async def cup_matches_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    state = load_cup_state()
    if not state.get("rounds"):
        await update.message.reply_text("لا توجد كأس محفوظة حاليًا.\nلبدء كأس جديد: /بدء_الكاس 5")
        return

    pending = cup_pending_round_key(state)
    if not pending:
        await update.message.reply_text(cup_results_text(state))
        return

    path = create_cup_matches_image(state, pending)
    if path:
        caption = cup_results_text(state) if pending == "champion" else f"🏆 مواجهات كأس المصيف — {cup_round_name(pending)}"
        await send_photo_path(update, path, caption)
    else:
        await update.message.reply_text(cup_results_text(state))

async def cup_after_fantasy_results(update: Update, day):
    state = load_cup_state()
    if not state.get("active"):
        return
    try:
        day_i = int(day)
        if day_i < int(state.get("start_day")) or day_i > int(state.get("end_day")):
            return
    except Exception:
        return

    state, round_key, msg = cup_process_day(day_i)
    if not round_key:
        return

    # أرسل نتائج المرحلة النصية
    await update.message.reply_text(f"🏆 كأس المصيف\n{msg}\n\n{cup_results_text(state)}")

    # بعد حساب اليوم: أرسل تصميم المرحلة القادمة، أو البطل إذا انتهى النهائي
    next_key = cup_pending_round_key(state)
    if next_key:
        path = create_cup_matches_image(state, next_key)
        if path:
            if next_key == "champion":
                caption = f"🏆 بطل كأس المصيف: {state.get('champion')}"
            else:
                caption = f"🏆 مواجهات {cup_round_name(next_key)} في كأس المصيف"
            await send_photo_path(update, path, caption)


# ============================================================
# V20 — اعتماد النتائج النهائي + إعادة الكأس من يوم
# ============================================================

def day_has_result_lines(message_text):
    lines = [l.strip() for l in (message_text or "").splitlines()[1:] if l.strip()]
    return bool(lines)

def clear_cup_match_result(match):
    match["score_a"] = None
    match["score_b"] = None
    match["active_a"] = None
    match["active_b"] = None
    match["winner"] = None
    match["note"] = ""
    return match

def reset_cup_from_day(state, from_day):
    if not state.get("rounds"):
        return state, "لا توجد بطولة كأس محفوظة."

    start_day = int(state.get("start_day", 0))
    end_day = int(state.get("end_day", start_day + 3))
    from_day = int(from_day)

    if from_day < start_day or from_day > end_day:
        return state, f"اليوم {from_day} خارج فترة الكأس الحالية."

    offset = from_day - start_day
    keys = [x[0] for x in CUP_ROUNDS]
    reset_key = keys[offset]

    # إذا نعيد من ربع/نصف/نهائي، نحتاج نضمن أن المرحلة مبنية من نتائج المرحلة السابقة.
    if reset_key == "qf":
        state["rounds"]["qf"]["matches"] = cup_build_next_round_matches(state, "r12")
    elif reset_key == "sf":
        state["rounds"]["sf"]["matches"] = cup_build_next_round_matches(state, "qf")
    elif reset_key == "final":
        state["rounds"]["final"]["matches"] = cup_build_next_round_matches(state, "sf")

    # امسح نتائج المرحلة المطلوبة وما بعدها، وامسح مواجهات المراحل التي بعدها
    for idx, key in enumerate(keys):
        rd = state["rounds"].get(key, {})
        if idx < offset:
            continue

        if idx == offset:
            rd["matches"] = [clear_cup_match_result(m) for m in rd.get("matches", [])]
            rd["completed"] = False
        else:
            rd["matches"] = []
            rd["completed"] = False

        state["rounds"][key] = rd

    state["champion"] = None
    state["active"] = True
    state["reset_from_day"] = from_day
    state["reset_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return state, f"تمت إعادة الكأس من اليوم {from_day} ✅"

async def approve_results_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    day = get_day(text)

    if is_locked(day):
        await update.message.reply_text(f"اليوم {day} مقفل ✅\nإذا تبي تعيد اعتماده اكتب أولًا:\n/فتح_يوم {day}")
        return

    if not os.path.exists(excel_file(day)):
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}. أضف التشكيلات أولًا.")
        return

    # لو أرسل نتائج داخل أمر الاعتماد، نعيد الحساب قبل الاعتماد.
    if day_has_result_lines(text):
        goals_count, clean_sheets = parse_results(text)
        goal_missing, clean_missing = validate_results_names(day, goals_count, clean_sheets)
        backup_files(f"before_approve_results_day_{day}", files=[excel_file(day), LOCKED_FILE, CUP_FILE])
        file_name = calculate_points(day, goals_count, clean_sheets)
    else:
        goal_missing, clean_missing = [], []
        file_name = excel_file(day)

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
        f"تم اعتماد نتائج اليوم {day} رسميًا ✅\n"
        f"تم قفل اليوم {day} 🔒\n\n"
        f"🏆 أسطورة اليوم: {legends_text} — {max_score} نقطة"
    )
    if warnings:
        caption += "\n\n" + "\n\n".join(warnings)

    # أرسل ملف اليوم المعتمد
    with open(file_name, "rb") as file:
        await update.message.reply_document(document=file, filename=file_name, caption=caption)

    # احسب الكأس فقط هنا، وليس في /نتائج
    try:
        await cup_after_fantasy_results(update, day)
    except Exception as e:
        await update.message.reply_text(f"تم اعتماد اليوم، لكن تعذر تحديث الكأس ❌\n{e}")

    # اقفل اليوم بعد الاعتماد
    days = load_locked_days()
    days.add(str(day))
    save_locked_days(days)

async def reset_cup_from_day_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nums = get_numbers(update.message.text)
    if not nums:
        await update.message.reply_text("اكتبها كذا:\n/إعادة_الكاس_من 7")
        return

    from_day = int(nums[0])
    state = load_cup_state()
    if not state.get("rounds"):
        await update.message.reply_text("لا توجد كأس محفوظة حاليًا.\nلبدء كأس جديد: /بدء_الكاس 7")
        return

    backup_files(f"before_reset_cup_from_{from_day}", files=[CUP_FILE] if os.path.exists(CUP_FILE) else [])
    state, msg = reset_cup_from_day(state, from_day)
    save_cup_state(state)

    pending = cup_pending_round_key(state)
    if pending:
        path = create_cup_matches_image(state, pending)
        if path:
            await send_photo_path(update, path, msg + f"\n\nالمرحلة الحالية: {cup_round_name(pending)}")
            return

    await update.message.reply_text(msg)

# -------------------- V18: قراءة النموذج الرسمي الطويل في /اضافه --------------------

def clean_pick_value(value):
    """
    تنظيف بسيط للاسم بدون Alias:
    - يشيل الإيموجي والزخارف
    - يحافظ على النص كما كتبه المستخدم قدر الإمكان
    """
    value = normalize_name(value)
    value = value.replace("：", ":")
    # إزالة رموز شائعة تظهر بعد أسماء اللاعبين
    value = re.sub(r"[👑🐐🔥⚽️✅🧤🏆⭐🌟🥇🥈🥉📋\-–—]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value

def extract_after_colon(line):
    line = str(line or "").strip().replace("：", ":")
    if ":" in line:
        return clean_pick_value(line.split(":", 1)[1])
    # احتياط لو كتب بدون نقطتين: "الحارس نايلاند"
    for key in ["الحارس", "حارس", "الكابتن", "كابتن", "اللاعب", "لاعب"]:
        if key in line:
            return clean_pick_value(line.split(key, 1)[1])
    return ""

def participant_line_name(line):
    line = normalize_name(line)
    if not line:
        return ""
    cleaned = re.sub(r"[^\u0600-\u06FFa-zA-Z\s]", " ", line)
    cleaned = normalize_name(cleaned)
    for p in PARTICIPANTS:
        if cleaned == p:
            return p
    return ""

def parse_pipe_lineup_lines(lines):
    data = {}
    bad_lines = []
    for line in lines:
        raw = line.strip()
        if not raw:
            continue
        if "|" not in raw:
            continue
        parts = [clean_pick_value(p) for p in raw.split("|")]
        if len(parts) != 6:
            bad_lines.append(raw)
            continue
        participant = normalize_name(parts[0])
        data[participant] = [normalize_name(x) for x in parts[1:]]
    return data, bad_lines

def parse_official_lineup_blocks(lines):
    """
    يقرأ النموذج الرسمي مثل:
    فهد فارس
    🧤 الحارس: أوناي سيمون
    اللاعب 1: أولمو
    اللاعب 2: سالم الدوسري
    اللاعب 3: داروين نونيز
    👑 الكابتن : نونيز

    ويدعم:
    لاعب: ميسي
    لاعب ٢: هالاند
    لاعب ٣: ...
    """
    data = {}
    bad_blocks = []
    current = None
    values = None

    def finish_current():
        nonlocal current, values
        if not current:
            return
        keeper = values.get("keeper", "")
        p1 = values.get("p1", "")
        p2 = values.get("p2", "")
        p3 = values.get("p3", "")
        captain = values.get("captain", "")
        if any([keeper, p1, p2, p3, captain]):
            # إذا ناقص شيء نخليه فارغ بدل ما نرفض كامل المشاركة
            data[current] = [keeper, p1, p2, p3, captain]
        current = None
        values = None

    def put_player(value, preferred=None):
        if not value:
            return
        if preferred and not values.get(preferred):
            values[preferred] = value
            return
        for key in ["p1", "p2", "p3"]:
            if not values.get(key):
                values[key] = value
                return

    for raw in lines:
        line = normalize_name(raw)
        if not line:
            continue

        # إذا لقى اسم مشارك جديد، يقفل السابق
        p_name = participant_line_name(line)
        if p_name:
            finish_current()
            current = p_name
            values = {"keeper": "", "p1": "", "p2": "", "p3": "", "captain": ""}
            continue

        if not current:
            continue

        line_no_space = line.replace(" ", "")
        value = extract_after_colon(line)

        if "الحارس" in line or line.startswith("حارس") or "🧤" in raw:
            if value:
                values["keeper"] = value
            continue

        if "الكابتن" in line or "كابتن" in line or "👑" in raw:
            if value:
                values["captain"] = value
            continue

        if "لاعب" in line:
            if not value:
                continue
            preferred = None
            if re.search(r"(3|٣|الثالث|ثالث)", line_no_space):
                preferred = "p3"
            elif re.search(r"(2|٢|الثاني|ثاني)", line_no_space):
                preferred = "p2"
            elif re.search(r"(1|١|الأول|اول|الأول|الاول)", line_no_space):
                preferred = "p1"
            put_player(value, preferred)
            continue

    finish_current()
    return data, bad_blocks

def parse_add_day_text(message_text):
    """
    يقبل الصيغتين:
    1) فهد|حارس|لاعب1|لاعب2|لاعب3|كابتن
    2) النموذج الرسمي الطويل لكل مشارك
    """
    lines = (message_text or "").splitlines()[1:]

    pipe_data, pipe_bad = parse_pipe_lineup_lines(lines)
    official_data, official_bad = parse_official_lineup_blocks(lines)

    # لو نفس المشارك موجود بالصيغتين، آخر قراءة من النموذج الرسمي تغطي
    data = {}
    data.update(pipe_data)
    data.update(official_data)

    # الأسطر السيئة فقط للبايب؛ لا نحاسب النموذج الطويل على الزخارف والعناوين
    bad_lines = pipe_bad[:]

    return data, bad_lines

async def add_day(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    day = get_day(text)

    if is_locked(day):
        await update.message.reply_text(f"اليوم {day} مقفل ✅\nلفتحه اكتب: /فتح_يوم {day}")
        return

    data, bad_lines = parse_add_day_text(text)

    if not data:
        await update.message.reply_text(
            "ما لقيت مشاركين بصيغة صحيحة.\n\n"
            "الصيغة المختصرة:\n"
            "/اضافه 5\n"
            "فهد فارس|الحارس|لاعب1|لاعب2|لاعب3|الكابتن\n\n"
            "أو النموذج الرسمي:\n"
            "فهد فارس\n"
            "🧤 الحارس: أوناي سيمون\n"
            "اللاعب 1: أولمو\n"
            "اللاعب 2: سالم الدوسري\n"
            "اللاعب 3: داروين نونيز\n"
            "👑 الكابتن : نونيز"
        )
        return

    backup_files(f"before_add_day_{day}", files=[excel_file(day), LOCKED_FILE])
    file_name, unknown = update_day_data(day, data)

    caption = f"تم إنشاء/تحديث ملف اليوم {day} ✅\nعدد المشاركين المرسلين: {len(data)}"
    caption += "\n\n✅ تم قبول الصيغتين: السطر الواحد + النموذج الرسمي"

    if unknown:
        caption += "\n\n⚠️ أسماء مشاركين غير موجودة بالقائمة:\n" + "\n".join(unknown)
    if bad_lines:
        caption += "\n\n⚠️ أسطر مختصرة لم أفهمها:\n" + "\n".join(bad_lines[:5])

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

    backup_files(f"before_temp_results_day_{day}", files=[excel_file(day), LOCKED_FILE, CUP_FILE])
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
        f"تم تحديث نقاط اليوم {day} مؤقتًا ✅\n"
        f"لن يتم حساب الكأس أو اعتماد اليوم إلا بعد الأمر:\n"
        f"/اعتماد_نتائج {day}\n\n"
        f"الأهداف:\n{goals_text}\n\n"
        f"الكلين شيت:\n{clean_text}\n\n"
        f"🏆 المتصدر/أسطورة اليوم حاليًا: {legends_text} — {max_score} نقطة"
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



# ============================================================
# V8 - الصور والتصاميم والإعلانات والأعلام
# ============================================================

SETTINGS_FILE = "bot_settings.json"
GENERATED_DIR = "generated_images"
FLAGS_ZIP = "worldcup_2026_flags_pack.zip"
FLAGS_JSON = "flags_map.json"
FLAGS_DIR = os.path.join("assets", "flags")

RESULT_ANNOUNCEMENT_TEMPLATE = """أساطير اليوم {day_name} لفانتزي المصيف :
\"          🏆( {legends} ) 🏆

📌 يرجى مراجعة النقاط، وفي حال وجود أي خطأ في الاحتساب أو تسجيل النقاط فإن مدة الاعتراض أو طلب المراجعة هي 24 ساعة من وقت إرسال هذه الرسالة.
⏳ بعد انتهاء المدة تُعتمد النتائج بشكل نهائي ولا يقبل أي اعتراض أو تعديل لاحقاً.

بالتوفيق للجميع ⚽️✅🧤🏆"""

MATCHES_ANNOUNCEMENT_TEMPLATE = """🏆 فانتزي المصيف 2026  🏆
ً           🔥🔥🔥 مباريات اليوم ( {day_name} )  🔥🔥🔥🔥  

{matches_text}

📋 نموذج المشاركة الرسمي المعتمد
🏆 تشكيلة الفانتزي - اليوم ( {day_name} )
🧤 الحارس:
 اللاعب 1:
 اللاعب 2:
 اللاعب 3:
👑 الكابتن :"""

# نحاول دعم تشكيل العربية داخل الصور، ولو المكتبات غير موجودة نستمر بدون كراش.
try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter, features as PIL_FEATURES
except Exception:
    Image = ImageDraw = ImageFont = ImageFilter = PIL_FEATURES = None

try:
    PIL_RAQM = bool(PIL_FEATURES and PIL_FEATURES.check("raqm"))
except Exception:
    PIL_RAQM = False

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
except Exception:
    arabic_reshaper = None
    get_display = None


def ar_text(text):
    text = "" if text is None else str(text)
    if arabic_reshaper and get_display:
        try:
            return get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
    return text


def _safe_filename(name):
    name = normalize_name(name)
    name = re.sub(r"[^A-Za-z0-9_.\-\u0600-\u06FF]+", "_", name).strip("_")
    return name or "file"


def ensure_generated_dir():
    os.makedirs(GENERATED_DIR, exist_ok=True)


def load_settings():
    default = {"auto_images": True}
    if not os.path.exists(SETTINGS_FILE):
        return default
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        default.update(data if isinstance(data, dict) else {})
    except Exception:
        pass
    return default


def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def admin_id_set():
    raw = os.getenv("ADMIN_IDS") or os.getenv("ADMINS") or ""
    ids = set()
    for part in raw.replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        try:
            ids.add(int(part))
        except Exception:
            pass
    return ids


def is_admin_user(update):
    ids = admin_id_set()
    if not ids:
        # إذا ما ضبطت ADMIN_IDS في Railway نخلي البوت يشتغل عشان ما ينقفل عليك.
        return True
    user = getattr(update, "effective_user", None)
    return bool(user and user.id in ids)


def admin_only(func):
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
        if not is_admin_user(update):
            await update.message.reply_text("هذا الأمر للمشرفين فقط 🔒")
            return
        return await func(update, context)
    return wrapper


async def who_am_i(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    chat = update.effective_chat
    await update.message.reply_text(
        f"User ID: {user.id if user else '-'}\n"
        f"Chat ID: {chat.id if chat else '-'}\n\n"
        "حط User ID في Railway داخل ADMIN_IDS لو تبي تقفل أوامر الإدارة."
    )


def font_candidates():
    """
    ترتيب خطوط عربي محسّن للتصاميم.
    البوت يختار أول خط موجود من القائمة، لذلك ارفع هذه الملفات بجانب bot.py:
    Tajawal-ExtraBold.ttf / Tajawal-Black.ttf / Cairo-Bold.ttf / NotoNaskhArabic-Bold.ttf / Amiri-Bold.ttf
    """
    names = [
        "Tajawal-ExtraBold.ttf",
        "Tajawal-Black.ttf",
        "Tajawal-Bold.ttf",
        "Cairo-Bold.ttf",
        "Cairo-Bold-1.ttf",
        "NotoNaskhArabic-Bold.ttf",
        "NotoNaskhArabic-Regular.ttf",
        "Amiri-Bold.ttf",
        "DejaVuSans-Bold.ttf",
    ]
    bases = ["", ".", "/app", "/mnt/data", "/usr/share/fonts/truetype/noto", "/usr/share/fonts/truetype/dejavu"]
    paths = []
    for base in bases:
        for name in names:
            paths.append(os.path.join(base, name) if base else name)
    return paths


def get_font(size):
    if not ImageFont:
        return None
    for path in font_candidates():
        try:
            if os.path.exists(path):
                return ImageFont.truetype(path, size)
        except Exception:
            pass
    try:
        return ImageFont.truetype("DejaVuSans-Bold.ttf", size)
    except Exception:
        return ImageFont.load_default()


def font_size(font, default=24):
    try:
        return int(getattr(font, "size", default) or default)
    except Exception:
        return default


def get_font_from_names(size, names):
    if not ImageFont:
        return None
    bases = ["", ".", "/app", "/mnt/data", "/usr/share/fonts/truetype/noto", "/usr/share/fonts/truetype/dejavu"]
    for base in bases:
        for name in names:
            path = os.path.join(base, name) if base else name
            try:
                if os.path.exists(path):
                    return ImageFont.truetype(path, size)
            except Exception:
                pass
    return get_font(size)


def get_arabic_fallback_font(size):
    """
    يستخدم فقط إذا ما كان RAQM متوفرًا.
    Noto Naskh يدعم العربية مع reshaper/bidi أفضل من بعض خطوط العناوين.
    """
    return get_font_from_names(size, [
        "NotoNaskhArabic-Bold.ttf",
        "NotoNaskhArabic-Regular.ttf",
        "Cairo-Bold.ttf",
        "Cairo-Bold-1.ttf",
        "Amiri-Bold.ttf",
        "Tajawal-Bold.ttf",
        "DejaVuSans-Bold.ttf",
    ])

def has_arabic(text):
    return bool(re.search(r"[\u0600-\u06FF]", str(text or "")))


def draw_text(draw, xy, text, font, fill="white", anchor="mm", align="center", max_width=None, spacing=8):
    """
    V12: رسم عربي آمن داخل الصور.
    - إذا RAQM متوفر: نرسم النص الأصلي direction=rtl حتى لا تظهر مربعات/فواصل داخل الكلمات.
    - إذا RAQM غير متوفر: نستخدم arabic_reshaper + bidi مع خط NotoNaskh كبديل.
    """
    text = "" if text is None else str(text)
    if font is None:
        font = get_font(24)

    if max_width:
        lines = wrap_text(draw, text, font, max_width)
        line_h = int(font_size(font, 24) * 1.30)
        total_h = line_h * len(lines)
        x, y = xy
        start_y = y - total_h / 2 if anchor == "mm" else y
        for i, line in enumerate(lines):
            draw_text(draw, (x, start_y + i * line_h), line, font, fill=fill, anchor="ma", align=align)
        return

    try:
        if has_arabic(text) and PIL_RAQM:
            draw.text(xy, text, font=font, fill=fill, anchor=anchor, align=align, direction="rtl", language="ar")
        elif has_arabic(text):
            f = get_arabic_fallback_font(font_size(font, 24))
            draw.text(xy, ar_text(text), font=f, fill=fill, anchor=anchor, align=align)
        else:
            draw.text(xy, text, font=font, fill=fill, anchor=anchor, align=align)
    except Exception:
        try:
            f = get_arabic_fallback_font(font_size(font, 24)) if has_arabic(text) else font
            display = ar_text(text) if has_arabic(text) else text
            draw.text(xy, display, font=f, fill=fill, anchor=anchor, align=align)
        except Exception:
            draw.text(xy, text, font=font, fill=fill, anchor=anchor)

def text_width(draw, text, font):
    text = "" if text is None else str(text)
    try:
        if has_arabic(text) and PIL_RAQM:
            bbox = draw.textbbox((0, 0), text, font=font, direction="rtl", language="ar")
        elif has_arabic(text):
            f = get_arabic_fallback_font(font_size(font, 24))
            bbox = draw.textbbox((0, 0), ar_text(text), font=f)
        else:
            bbox = draw.textbbox((0, 0), text, font=font)
        return bbox[2] - bbox[0]
    except Exception:
        return len(str(text)) * 12

def wrap_text(draw, text, font, max_width):
    text = "" if text is None else str(text)
    final_lines = []
    for raw_line in text.splitlines() or [""]:
        words = raw_line.split()
        if not words:
            final_lines.append("")
            continue
        line = ""
        for word in words:
            test = word if not line else line + " " + word
            if text_width(draw, test, font) <= max_width:
                line = test
            else:
                if line:
                    final_lines.append(line)
                line = word
        if line:
            final_lines.append(line)
    return final_lines


def rounded_rect(draw, box, radius=24, fill=None, outline=None, width=1):
    try:
        draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)
    except Exception:
        draw.rectangle(box, fill=fill, outline=outline, width=width)


def make_canvas(width, height, theme="purple"):
    if not Image:
        raise RuntimeError("Pillow غير مثبت. أضف Pillow في requirements.txt")
    img = Image.new("RGB", (width, height), "#101827")
    draw = ImageDraw.Draw(img)
    # خلفية متدرجة بسيطة
    for y in range(height):
        r = int(18 + y / height * 12)
        g = int(24 + y / height * 10)
        b = int(45 + y / height * 35)
        draw.line([(0, y), (width, y)], fill=(r, g, b))
    # دوائر ديكور
    for i, (x, y, rr, col) in enumerate([
        (120, 80, 180, "#7C3AED"), (width-130, 120, 220, "#2563EB"),
        (width//2, height-80, 300, "#F2B705")
    ]):
        overlay = Image.new("RGBA", (width, height), (0,0,0,0))
        od = ImageDraw.Draw(overlay)
        od.ellipse((x-rr, y-rr, x+rr, y+rr), fill=col + "30")
        overlay = overlay.filter(ImageFilter.GaussianBlur(40))
        img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
        draw = ImageDraw.Draw(img)
    return img, draw


def ensure_flags_assets():
    if os.path.exists(FLAGS_DIR) and os.path.exists(FLAGS_JSON):
        return
    if os.path.exists(FLAGS_ZIP):
        try:
            import zipfile
            with zipfile.ZipFile(FLAGS_ZIP, "r") as z:
                z.extractall(".")
        except Exception:
            pass


def load_flags_map():
    ensure_flags_assets()
    data = {}
    if os.path.exists(FLAGS_JSON):
        try:
            with open(FLAGS_JSON, "r", encoding="utf-8") as f:
                obj = json.load(f)
            aliases = obj.get("aliases", {}) if isinstance(obj, dict) else {}
            for k, v in aliases.items():
                data[normalize_name(k)] = v
                data[normalize_name(k).lower()] = v
        except Exception:
            pass
    return data


def flag_path_for(team_name):
    flags = load_flags_map()
    name = normalize_name(team_name)
    filename = flags.get(name) or flags.get(name.lower())
    if filename:
        path = os.path.join(FLAGS_DIR, filename)
        if os.path.exists(path):
            return path
    return None


def paste_flag(base, team_name, box):
    if not Image:
        return
    path = flag_path_for(team_name)
    x1, y1, x2, y2 = box
    w, h = int(x2-x1), int(y2-y1)
    if path and os.path.exists(path):
        try:
            flag = Image.open(path).convert("RGBA")
            flag.thumbnail((w, h), Image.LANCZOS)
            px = int(x1 + (w - flag.width) / 2)
            py = int(y1 + (h - flag.height) / 2)
            base.paste(flag, (px, py), flag)
            return
        except Exception:
            pass
    # بديل إذا العلم ناقص
    d = ImageDraw.Draw(base)
    rounded_rect(d, box, radius=18, fill="#FFFFFF22", outline="#FFFFFF55", width=2)
    f = get_font(24)
    draw_text(d, ((x1+x2)//2, (y1+y2)//2), normalize_name(team_name)[:2], f, fill="#FFFFFF")


def build_result_announcement(day, winners):
    day_name = ordinal_day(day)
    legends = " + ".join(winners) if winners else "لا يوجد"
    return RESULT_ANNOUNCEMENT_TEMPLATE.format(day_name=day_name, legends=legends)


def build_day_summary(day):
    rows = read_day_rows(day)
    participants = [r for r in rows if r.get("participated")]
    max_score = max([r["total"] for r in rows], default=0)
    winners = [r["participant"] for r in rows if r["total"] == max_score and max_score > 0]
    sorted_rows = sorted(rows, key=lambda r: r["total"], reverse=True)
    top_captain = max(rows, key=lambda r: r.get("captain_points", 0), default=None)
    top_keeper = max(rows, key=lambda r: r.get("keeper_points", 0), default=None)

    player_points = Counter()
    keeper_points = Counter()
    for r in rows:
        for k, pk in (("p1", "p1_points"), ("p2", "p2_points"), ("p3", "p3_points")):
            if r.get(pk, 0) > 0 and not is_no_participation(r.get(k)):
                player_points[r[k]] += r[pk]
        if r.get("keeper_points", 0) > 0 and not is_no_participation(r.get("keeper")):
            keeper_points[r["keeper"]] += r["keeper_points"]

    return {
        "rows": rows,
        "participants": participants,
        "max_score": max_score,
        "winners": winners,
        "sorted_rows": sorted_rows,
        "top_captain": top_captain,
        "top_keeper": top_keeper,
        "player_points": player_points,
        "keeper_points": keeper_points,
    }


def create_daily_result_image(day, goals_count=None, clean_sheets=None):
    ensure_generated_dir()
    data = build_day_summary(day)
    rows = data["sorted_rows"]
    h = max(1050, 470 + len(rows) * 58)
    img, draw = make_canvas(1400, h)
    title_f = get_font(58)
    sub_f = get_font(34)
    small_f = get_font(26)
    row_f = get_font(28)

    draw_text(draw, (700, 85), f"فانتزي المصيف 2026 — اليوم {ordinal_day(day)}", title_f, fill="#FFFFFF")
    draw_text(draw, (700, 145), "نتائج اليوم وترتيب المشاركين", sub_f, fill="#FDE68A")

    winners = data["winners"]
    legends = " + ".join(winners) if winners else "لا يوجد"
    rounded_rect(draw, (80, 190, 1320, 340), radius=38, fill="#7C3AEDDD", outline="#FFFFFF33", width=2)
    draw_text(draw, (700, 235), "🏆 أسطورة اليوم 🏆", sub_f, fill="#FFFFFF")
    draw_text(draw, (700, 295), f"{legends} — {data['max_score']} نقطة", get_font(44), fill="#FFF6D6")

    # كروت جانبية
    top_cap = data["top_captain"]
    top_keep = data["top_keeper"]
    cards = [
        (80, 370, 455, 500, "👑 أفضل كابتن", f"{top_cap['participant']} +{top_cap['captain_points']}" if top_cap else "-", "#2563EB"),
        (512, 370, 887, 500, "🧤 أفضل حارس", f"{top_keep['participant']} +{top_keep['keeper_points']}" if top_keep else "-", "#10B981"),
        (945, 370, 1320, 500, "👥 المشاركون", f"{len(data['participants'])} مشارك", "#F59E0B"),
    ]
    for x1,y1,x2,y2,t,v,c in cards:
        rounded_rect(draw, (x1,y1,x2,y2), radius=28, fill=c+"DD", outline="#FFFFFF33", width=2)
        draw_text(draw, ((x1+x2)//2, y1+38), t, small_f, fill="#FFFFFF")
        draw_text(draw, ((x1+x2)//2, y1+92), v, sub_f, fill="#FFFFFF")

    # جدول ترتيب اليوم
    y = 550
    rounded_rect(draw, (80, y, 1320, y+58), radius=18, fill="#FFFFFF22", outline="#FFFFFF30", width=1)
    headers = [(1180, "المشارك"), (850, "النقاط"), (640, "الحارس"), (410, "الكابتن"), (170, "المركز")]
    for x, t in headers:
        draw_text(draw, (x, y+30), t, small_f, fill="#FFFFFF")
    y += 70
    for idx, r in enumerate(rows, start=1):
        fill = "#FFFFFF16" if idx % 2 else "#FFFFFF0C"
        if idx == 1:
            fill = "#F2B70544"
        rounded_rect(draw, (80, y, 1320, y+48), radius=14, fill=fill, outline="#FFFFFF18", width=1)
        draw_text(draw, (1180, y+25), r["participant"], row_f, fill="#FFFFFF")
        draw_text(draw, (850, y+25), str(r["total"]), row_f, fill="#FDE68A")
        draw_text(draw, (640, y+25), f"+{r['keeper_points']}", row_f, fill="#A7F3D0")
        draw_text(draw, (410, y+25), f"+{r['captain_points']}", row_f, fill="#C4B5FD")
        medal = "🥇" if idx == 1 else "🥈" if idx == 2 else "🥉" if idx == 3 else str(idx)
        draw_text(draw, (170, y+25), medal, row_f, fill="#FFFFFF")
        y += 58

    # الهدافين والكلين شيت
    y += 25
    box_h = 180
    rounded_rect(draw, (80, y, 670, y+box_h), radius=26, fill="#111827CC", outline="#FFFFFF22", width=2)
    rounded_rect(draw, (730, y, 1320, y+box_h), radius=26, fill="#111827CC", outline="#FFFFFF22", width=2)
    goals_lines = []
    if goals_count:
        goals_lines = [f"{p} — {c}" for p,c in goals_count.items()]
    else:
        goals_lines = [f"{p} — {pts} نقطة" for p,pts in data["player_points"].most_common(5)] or ["لا يوجد"]
    clean_lines = clean_sheets or list(data["keeper_points"].keys()) or ["لا يوجد"]
    draw_text(draw, (375, y+35), "⚽ الهدافون", sub_f, fill="#FFFFFF")
    draw_text(draw, (375, y+105), "\n".join(goals_lines[:4]), small_f, fill="#E5E7EB", max_width=520)
    draw_text(draw, (1025, y+35), "🧤 الكلين شيت", sub_f, fill="#FFFFFF")
    draw_text(draw, (1025, y+105), "\n".join(clean_lines[:4]), small_f, fill="#E5E7EB", max_width=520)

    path = os.path.join(GENERATED_DIR, f"daily_result_day_{day}.png")
    img.save(path, quality=95)
    return path


def build_daily_summary_text(day):
    data = build_day_summary(day)
    winners = " + ".join(data["winners"]) if data["winners"] else "لا يوجد"
    top_rows = data["sorted_rows"][:5]
    lines = [
        f"📊 ملخص اليوم {ordinal_day(day)}",
        f"👥 عدد المشاركين: {len(data['participants'])}",
        f"🏆 أسطورة اليوم: {winners}",
        f"🔥 أعلى نقاط: {data['max_score']}",
        "",
        "أول 5 في اليوم:",
    ]
    for i, r in enumerate(top_rows, start=1):
        lines.append(f"{i}. {r['participant']} — {r['total']} نقطة")
    return "\n".join(lines)


def create_overall_ranking_image(start_day=1, end_day=31):
    """
    V21: صورة ترتيب الفانتزي بنفس هوية التصاميم الجديدة
    بدال الجدول القديم اللي كانت بعض الخانات تطلع بيضاء/باهتة.
    """
    ensure_generated_dir()
    stats = collect_stats(start_day, end_day)
    ranking = stats.get("ranking", [])
    count = max(len(ranking), 1)
    width = 1200

    # تصميم مرن حسب عدد المشاركين
    if count <= 8:
        row_h, gap, name_size = 92, 12, 30
    elif count <= 12:
        row_h, gap, name_size = 78, 10, 26
    else:
        row_h, gap, name_size = 66, 8, 22

    header_h = 58
    content_h = header_h + 24 + count * row_h + max(0, count - 1) * gap
    height = max(900, 245 + content_h + 175)

    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, "ترتيب فانتزي المصيف 2026", f"من اليوم {start_day} إلى اليوم {end_day}", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=112, accent="#22C55E")

    leader = ranking[0] if ranking else None
    leader_points = int(stats.get("totals", {}).get(leader, 0) or 0) if leader else 0

    y = fy1 + 34

    # هيدر الجدول
    rounded_rect(draw, (92, y, width-92, y+header_h), radius=20, fill="#05070D", outline="#FFFFFF70", width=1)
    draw_text(draw, (1080, y+header_h//2), "المركز", get_font(22), fill="#FDE68A")
    draw_text(draw, (850, y+header_h//2), "المشارك", get_font(22), fill="#FDE68A")
    draw_text(draw, (570, y+header_h//2), "النقاط", get_font(22), fill="#FDE68A")
    draw_text(draw, (385, y+header_h//2), "الفارق", get_font(22), fill="#FDE68A")
    draw_text(draw, (190, y+header_h//2), "أسطورة", get_font(22), fill="#FDE68A")
    y += header_h + 24

    current_rank = 0
    last_score = None
    real_index = 0
    medals = {1: "🥇", 2: "🥈", 3: "🥉"}

    for name in ranking:
        real_index += 1
        score = int(stats.get("totals", {}).get(name, 0) or 0)
        if score != last_score:
            current_rank = real_index
            last_score = score

        accent = "#F59E0B" if current_rank == 1 else v16_accent(real_index)
        fill = "#1A1407" if current_rank == 1 else "#0B1020"
        rounded_rect(draw, (92, y, width-92, y+row_h), radius=22, fill=fill, outline=accent, width=2)
        cy = y + row_h//2

        rank_label = medals.get(current_rank, str(current_rank))
        rank_color = "#FDE68A" if current_rank <= 3 else "#FFFFFF"
        draw_text(draw, (1080, cy), rank_label, get_font(max(24, name_size)), fill=rank_color)

        # الاسم أوضح وبمساحة كافية
        draw_text(draw, (835, cy), name, get_font(name_size), fill="#FFFFFF", max_width=330)

        # النقاط
        draw_text(draw, (570, cy), str(score), get_font(max(26, name_size+2)), fill="#FDE68A")

        # الفارق عن المتصدر
        diff = leader_points - score
        diff_text = "—" if diff == 0 else str(diff)
        draw_text(draw, (385, cy), diff_text, get_font(max(24, name_size)), fill="#E5E7EB")

        # مرات أسطورة اليوم
        legends = int(stats.get("daily_wins", {}).get(name, 0) or 0)
        legends_text = str(legends)
        draw_text(draw, (190, cy), legends_text, get_font(max(24, name_size)), fill="#C4B5FD")

        y += row_h + gap

    footer_event(draw, width, height)

    path = os.path.join(GENERATED_DIR, f"overall_{start_day}_{end_day}.png")
    img.save(path, quality=95)
    return path

def create_legends_image(start_day=1, end_day=31):
    ensure_generated_dir()
    stats = collect_stats(start_day, end_day)
    days = stats["days"]
    h = max(800, 260 + len(days) * 72)
    img, draw = make_canvas(1400, h)
    draw_text(draw, (700, 90), "سجل أساطير الفانتزي", get_font(60), fill="#FFFFFF")
    draw_text(draw, (700, 150), f"من اليوم {start_day} إلى اليوم {end_day}", get_font(30), fill="#FDE68A")
    y = 230
    for day in days:
        info = stats["per_day"].get(day, {})
        winners = info.get("winners", [])
        names = " + ".join(winners) if winners else "لا يوجد"
        max_score = info.get("max_score", 0)
        rounded_rect(draw, (110, y, 1290, y+56), radius=18, fill="#FFFFFF14", outline="#FFFFFF22", width=1)
        draw_text(draw, (1130, y+29), f"اليوم {ordinal_day(day)}", get_font(30), fill="#FFFFFF")
        draw_text(draw, (700, y+29), names, get_font(30), fill="#FDE68A")
        draw_text(draw, (210, y+29), f"{max_score} نقطة", get_font(28), fill="#A7F3D0")
        y += 72
    path = os.path.join(GENERATED_DIR, f"legends_{start_day}_{end_day}.png")
    img.save(path, quality=95)
    return path


def parse_range_from_text(text):
    nums = get_numbers(text or "")
    if len(nums) >= 2:
        a, b = nums[0], nums[1]
        return (min(a, b), max(a, b))
    if len(nums) == 1:
        return (1, nums[0])
    return (1, 31)


def page_name_from_command(text, default="لوحة عامة"):
    text = text or ""
    # نحذف الأمر والأرقام ونبقي اسم الصفحة
    parts = text.splitlines()[0].split()
    if not parts:
        return default
    rest = []
    for p in parts[1:]:
        if re.fullmatch(r"\d+", p):
            continue
        rest.append(p)
    name = " ".join(rest).strip()
    return name or default


def find_sheet_name(wb, requested):
    requested = normalize_name(requested)
    if requested in wb.sheetnames:
        return requested
    for s in wb.sheetnames:
        if requested and requested in s:
            return s
    for s in wb.sheetnames:
        if s in requested:
            return s
    return wb.sheetnames[0] if wb.sheetnames else None


def render_worksheet_to_images(xlsx_path, sheet_name, max_rows_per_image=32):
    ensure_generated_dir()
    wb = load_workbook(xlsx_path, data_only=False)
    sheet_name = find_sheet_name(wb, sheet_name)
    if not sheet_name:
        return []
    ws = wb[sheet_name]

    # نحدد حدود البيانات المفيدة بدون الأعمدة المخفية قدر الإمكان
    max_col = min(ws.max_column, 12)
    non_empty_rows = []
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        if any(v not in (None, "") for v in vals):
            non_empty_rows.append(r)
    if not non_empty_rows:
        non_empty_rows = [1]
    min_row, max_row = min(non_empty_rows), max(non_empty_rows)

    images = []
    col_w = 170
    row_h = 54
    title_h = 95
    page_index = 1
    for start in range(min_row, max_row + 1, max_rows_per_image):
        end = min(start + max_rows_per_image - 1, max_row)
        width = max(1100, max_col * col_w + 100)
        height = title_h + (end-start+1) * row_h + 80
        img, draw = make_canvas(width, height)
        draw_text(draw, (width//2, 52), f"{sheet_name}" + (f" — {page_index}" if max_row-min_row+1 > max_rows_per_image else ""), get_font(42), fill="#FFFFFF")
        y = title_h
        for r in range(start, end + 1):
            x = 50
            base_fill = "#FFFFFF14" if r % 2 else "#FFFFFF0A"
            for c in range(1, max_col + 1):
                cell = ws.cell(r, c)
                val = "" if cell.value is None else str(cell.value)
                fill = base_fill
                try:
                    fg = cell.fill.fgColor.rgb
                    if fg and fg != "00000000":
                        fill = "#" + fg[-6:] + "DD"
                except Exception:
                    pass
                rounded_rect(draw, (x, y, x+col_w-6, y+row_h-6), radius=10, fill=fill, outline="#FFFFFF20", width=1)
                font = get_font(22 if r != 1 else 24)
                draw_text(draw, (x + (col_w-6)//2, y + (row_h-6)//2), val[:42], font, fill="#FFFFFF", max_width=col_w-20)
                x += col_w
            y += row_h
        out = os.path.join(GENERATED_DIR, f"sheet_{_safe_filename(sheet_name)}_{page_index}.png")
        img.save(out, quality=92)
        images.append(out)
        page_index += 1
    return images


def parse_matches_text(text):
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    if not lines:
        return "اليوم", []
    # أول سطر هو الأمر
    if len(lines) == 1:
        return "اليوم", []
    day_name = lines[1]
    matches = []
    for line in lines[2:]:
        # يدعم: فريق|فريق|وقت أو فريق × فريق | وقت
        if "|" in line:
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 3:
                matches.append((parts[0], parts[1], parts[2]))
            elif len(parts) == 2 and "×" in parts[0]:
                a, b = [p.strip() for p in parts[0].split("×", 1)]
                matches.append((a, b, parts[1]))
        elif "×" in line:
            before, *after = re.split(r"[-—|]", line, maxsplit=1)
            a, b = [p.strip() for p in before.split("×", 1)]
            time = after[0].strip() if after else ""
            matches.append((a, b, time))
    return day_name, matches


def build_matches_announcement(day_name, matches):
    matches_text = "\n".join([f"{a} × {b} — {t}" for a,b,t in matches]) or ""
    return MATCHES_ANNOUNCEMENT_TEMPLATE.format(day_name=day_name, matches_text=matches_text)


def create_matches_image(day_name, matches):
    """تصميم مباريات للقروب: صورة أفقية واضحة بالأعلام، بدون قالب جامد.
    يضبط حجم الصفوف حسب عدد المباريات ويقلل الفراغ.
    """
    ensure_generated_dir()
    count = max(len(matches), 1)
    width = 1600
    # ارتفاع مناسب للقروب، ويزيد إذا المباريات كثيرة
    row_h = 170 if count <= 3 else 140
    gap = 22 if count <= 3 else 16
    header_h = 230
    footer_h = 125
    height = max(900, header_h + count * row_h + max(0, count-1) * gap + footer_h + 40)

    img, draw = make_canvas(width, height)

    # طبقة تظليل في المنتصف عشان النص يبان
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle((60, 45, width-60, height-45), radius=48, fill="#02061770", outline="#FFFFFF18", width=2)
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)

    # العنوان
    draw_text(draw, (width//2, 88), "مونديال المصيف 2026", get_font(58), fill="#FFFFFF")
    draw_text(draw, (width//2, 154), f"مباريات اليوم {day_name}", get_font(44), fill="#FDE68A")
    draw.line((210, 205, width-210, 205), fill="#FFFFFF35", width=2)

    colors = ["#7C3AED", "#2563EB", "#0891B2", "#059669", "#D97706", "#DC2626", "#4F46E5"]
    y = header_h
    for i, (a, b, t) in enumerate(matches, start=1):
        c = colors[(i-1) % len(colors)]
        # كرت المباراة
        rounded_rect(draw, (90, y, width-90, y+row_h), radius=36, fill=c+"D5", outline="#FFFFFF30", width=2)

        # مناطق الأعلام
        flag_w = 180 if count <= 3 else 145
        flag_h = 110 if count <= 3 else 90
        cy = y + row_h//2
        # يمين الفريق الأول
        paste_flag(img, a, (width-285, cy-flag_h//2, width-105, cy+flag_h//2))
        # يسار الفريق الثاني
        paste_flag(img, b, (105, cy-flag_h//2, 285, cy+flag_h//2))

        # أسماء المنتخبات
        name_font = get_font(46 if count <= 3 else 38)
        draw_text(draw, (width-470, cy), a, name_font, fill="#FFFFFF", max_width=360)
        draw_text(draw, (470, cy), b, name_font, fill="#FFFFFF", max_width=360)

        # علامة × والوقت
        draw_text(draw, (width//2, cy-20), "×", get_font(58 if count <= 3 else 48), fill="#FDE68A")
        if t:
            badge_w = 210 if count <= 3 else 180
            badge_h = 44 if count <= 3 else 38
            rounded_rect(draw, (width//2-badge_w//2, cy+32, width//2+badge_w//2, cy+32+badge_h), radius=18, fill="#020617E6", outline="#FFFFFF40", width=1)
            draw_text(draw, (width//2, cy+32+badge_h//2), t, get_font(26 if count <= 3 else 22), fill="#FFFFFF")
        y += row_h + gap

    # عبارة الختام
    footer_y = min(height-86, y+55)
    draw_text(draw, (width//2, footer_y), "حياكم في محلكم 🏆", get_font(42), fill="#FFFFFF")

    path = os.path.join(GENERATED_DIR, f"matches_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path


async def send_photo_path(update, path, caption=None):
    with open(path, "rb") as f:
        await update.message.reply_photo(photo=f, caption=caption or "")


# -------------------- أوامر الصور الجديدة --------------------

async def enable_auto_images(update: Update, context: ContextTypes.DEFAULT_TYPE):
    settings = load_settings()
    settings["auto_images"] = True
    save_settings(settings)
    await update.message.reply_text("تم تفعيل الصور التلقائية ✅")


async def disable_auto_images(update: Update, context: ContextTypes.DEFAULT_TYPE):
    settings = load_settings()
    settings["auto_images"] = False
    save_settings(settings)
    await update.message.reply_text("تم إيقاف الصور التلقائية ✅")


async def daily_image_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    if not os.path.exists(excel_file(day)):
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}.")
        return
    try:
        path = create_daily_result_image(day)
        await send_photo_path(update, path, build_result_announcement(day, build_day_summary(day)["winners"]))
    except Exception as e:
        await update.message.reply_text(f"تعذر إنشاء صورة اليوم ❌\n{e}")


async def overall_image_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    start_day, end_day = parse_range_from_text(update.message.text)
    try:
        path = create_overall_ranking_image(start_day, end_day)
        await send_photo_path(update, path, f"الترتيب العام من اليوم {start_day} إلى {end_day} ✅")
    except Exception as e:
        await update.message.reply_text(f"تعذر إنشاء صورة الترتيب ❌\n{e}")


async def legends_image_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    start_day, end_day = parse_range_from_text(update.message.text)
    try:
        path = create_legends_image(start_day, end_day)
        await send_photo_path(update, path, f"سجل الأساطير من اليوم {start_day} إلى {end_day} ✅")
    except Exception as e:
        await update.message.reply_text(f"تعذر إنشاء صورة الأساطير ❌\n{e}")


async def dashboard_sheet_image_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    start_day, end_day = parse_range_from_text(update.message.text)
    sheet = page_name_from_command(update.message.text, "لوحة عامة")
    await update.message.reply_text("جاري تجهيز صورة الصفحة... ⏳")
    try:
        xlsx, stats = create_dashboard(start_day, end_day)
        if not stats.get("days"):
            await update.message.reply_text("ما فيه أيام في النطاق المطلوب.")
            return
        images = render_worksheet_to_images(xlsx, sheet)
        if not images:
            await update.message.reply_text("ما قدرت أطلع صورة للصفحة.")
            return
        for i, path in enumerate(images, start=1):
            cap = f"{sheet}" if len(images) == 1 else f"{sheet} — جزء {i}"
            await send_photo_path(update, path, cap)
    except Exception as e:
        await update.message.reply_text(f"تعذر إنشاء صورة الإحصائيات ❌\n{e}")


async def all_dashboard_images_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    start_day, end_day = parse_range_from_text(update.message.text)
    await update.message.reply_text("جاري تجهيز صور كل الصفحات... ⏳")
    try:
        xlsx, stats = create_dashboard(start_day, end_day)
        if not stats.get("days"):
            await update.message.reply_text("ما فيه أيام في النطاق المطلوب.")
            return
        wb = load_workbook(xlsx, read_only=True)
        for sheet in wb.sheetnames:
            images = render_worksheet_to_images(xlsx, sheet)
            for i, path in enumerate(images, start=1):
                cap = f"{sheet}" if len(images) == 1 else f"{sheet} — جزء {i}"
                await send_photo_path(update, path, cap)
        wb.close()
    except Exception as e:
        await update.message.reply_text(f"تعذر إنشاء صور الإحصائيات ❌\n{e}")


async def announcement_day_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    if not os.path.exists(excel_file(day)):
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}.")
        return
    await update.message.reply_text(build_result_announcement(day, build_day_summary(day)["winners"]))


async def summary_day_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day = get_day(update.message.text)
    if not os.path.exists(excel_file(day)):
        await update.message.reply_text(f"ما لقيت ملف اليوم {day}.")
        return
    await update.message.reply_text(build_daily_summary_text(day))


async def matches_design_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day_name, matches = parse_matches_text(update.message.text)
    if not matches:
        await update.message.reply_text(
            "اكتبها كذا:\n"
            "/تصميم_مباريات\n"
            "السادس\n"
            "فرنسا|البرازيل|10:00 م\n"
            "الأرجنتين|ألمانيا|12:00 ص"
        )
        return
    try:
        path = create_matches_image(day_name, matches)
        caption = build_matches_announcement(day_name, matches)
        await send_photo_path(update, path, caption)
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم المباريات ❌\n{e}")


async def clean_temp_files(update: Update, context: ContextTypes.DEFAULT_TYPE):
    removed = 0
    targets = [GENERATED_DIR, "imports", "uploads", "restore_uploads"]
    for folder in targets:
        if not os.path.isdir(folder):
            continue
        for name in os.listdir(folder):
            path = os.path.join(folder, name)
            try:
                if os.path.isfile(path):
                    os.remove(path)
                    removed += 1
            except Exception:
                pass
    await update.message.reply_text(f"تم تنظيف الملفات المؤقتة ✅\nعدد الملفات المحذوفة: {removed}")


# -------------------- إعادة تعريف /start و /نتائج لإضافة الصور التلقائية --------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "الأوامر الأساسية:\n"
        "/اضافه 5\n"
        "/نتائج 5\n"
        "/احصائيات\n"
        "/احصائيات 1 6\n"
        "/ترتيب_نص\n\n"
        "أوامر الصور:\n"
        "/صورة_اليوم 6\n"
        "/صورة_الترتيب 1 6\n"
        "/صورة_الاساطير 1 6\n"
        "/صورة_احصائيات 1 6 لوحة عامة\n"
        "/صور_الاحصائيات 1 6\n"
        "/تفعيل_الصور_التلقائية\n"
        "/إيقاف_الصور_التلقائية\n\n"
        "أوامر المباريات:\n"
        "/تصميم_مباريات\n\n"
        "أوامر الفحص والنشر:\n"
        "/الأيام\n"
        "/فحص 5\n"
        "/مشاركين 5\n"
        "/اسطورة 5\n"
        "/مقارنة 4 5\n"
        "/اعلان_اليوم 5\n"
        "/ملخص_اليوم 5\n\n"
        "أوامر الاستيراد والنسخ:\n"
        "/استيراد_ملف — أرسل ملف الإكسل لحاله ثم اكتب الأمر\n"
        "/اعتماد_استيراد\n"
        "/إلغاء_استيراد\n"
        "/نسخة_احتياطية\n"
        "/استرجاع_نسخة — أرسل ملف ZIP لحاله ثم اكتب الأمر\n"
        "/تنظيف_الأيام\n"
        "/تنظيف_الملفات\n\n"
        "أوامر الأمان:\n"
        "/مسح_نتائج 5\n"
        "/مسح_يوم 5\n"
        "/مسح_الكل تأكيد\n"
        "/استرجاع_آخر\n"
        "/قفل_يوم 5\n"
        "/فتح_يوم 5\n"
        "/من_انا"
    )


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

    if load_settings().get("auto_images", True):
        try:
            path = create_daily_result_image(day, goals_count=goals_count, clean_sheets=clean_sheets)
            await send_photo_path(update, path, build_result_announcement(day, winners))
        except Exception as e:
            await update.message.reply_text(f"تم حساب النتائج، لكن تعذر إنشاء الصورة التلقائية ❌\n{e}")



# ============================================================
# V10 - إضافات نهائية: مواجهات بعد النتائج، كأس الفانتزي، بطاقات، تقارير، هدافين، نتائج مباريات
# ============================================================
import random

MATCHUPS_FILE = "matchups_history.json"
CUP_FILE = "fantasy_cup_history.json"


def font_candidates():
    """خطوط عربية مرتبة بالأولوية: عنوان ثم نص، يختار أول الموجود."""
    return [
        "Tajawal-Black.ttf",
        "Tajawal-ExtraBold.ttf",
        "Cairo-Bold.ttf",
        "Cairo-Bold-1.ttf",
        "NotoNaskhArabic-Bold.ttf",
        "Amiri-Bold.ttf",
        "NotoNaskhArabic-Regular.ttf",
        "/app/Tajawal-Black.ttf",
        "/app/Tajawal-ExtraBold.ttf",
        "/app/Cairo-Bold.ttf",
        "/app/Cairo-Bold-1.ttf",
        "/app/NotoNaskhArabic-Bold.ttf",
        "/app/Amiri-Bold.ttf",
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]


def load_json_file(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if data is not None else default
    except Exception:
        return default


def save_json_file(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def active_participants_for_day(day):
    return [r for r in read_day_rows(day) if r.get("participated")]


def daily_score_map(day):
    rows = read_day_rows(day)
    return {r["participant"]: r["total"] for r in rows}


def previous_ranking_before_day(day):
    try:
        d = int(day)
    except Exception:
        d = 1
    if d <= 1:
        return {name: i for i, name in enumerate(PARTICIPANTS, start=1)}
    stats = collect_stats(1, d - 1)
    if not stats["days"]:
        return {name: i for i, name in enumerate(PARTICIPANTS, start=1)}
    return {name: i for i, name in enumerate(stats["ranking"], start=1)}


def generate_matchups_for_day(day, force=False):
    """مواجهات اليوم تظهر بعد النتائج فقط. يدخل فقط المشاركون فعليًا في ذلك اليوم."""
    day_key = str(day)
    history = load_json_file(MATCHUPS_FILE, {})
    if (not force) and day_key in history:
        return history[day_key]

    active = [r["participant"] for r in active_participants_for_day(day)]
    if not active:
        result = {"day": day_key, "pairs": [], "bye": None, "note": "لا يوجد مشاركون"}
        history[day_key] = result
        save_json_file(MATCHUPS_FILE, history)
        return result

    # قرعة ثابتة حسب اليوم، غير مرتبطة بالنقاط.
    names = sorted(active)
    rnd = random.Random(int(day) * 2026 + 77)
    rnd.shuffle(names)

    bye = None
    if len(names) % 2 == 1:
        bye = names.pop()

    pairs = []
    scores = daily_score_map(day)
    for i in range(0, len(names), 2):
        a, b = names[i], names[i + 1]
        pa, pb = scores.get(a, 0), scores.get(b, 0)
        if pa > pb:
            winner, status = a, "فوز"
        elif pb > pa:
            winner, status = b, "فوز"
        else:
            # في التعادل يفوز الأعلى ترتيبًا قبل الجولة
            ranks = previous_ranking_before_day(day)
            winner = a if ranks.get(a, 999) <= ranks.get(b, 999) else b
            status = "تعادل — حُسم بالأفضلية"
        pairs.append({"a": a, "a_points": pa, "b": b, "b_points": pb, "winner": winner, "status": status})

    result = {"day": day_key, "pairs": pairs, "bye": bye, "note": ""}
    history[day_key] = result
    save_json_file(MATCHUPS_FILE, history)
    return result


def matchup_lines(day):
    data = generate_matchups_for_day(day)
    if not data.get("pairs") and not data.get("bye"):
        return ["⚔️ مواجهات اليوم:", "لا توجد مواجهات — لا يوجد مشاركون في هذا اليوم."]
    lines = [f"⚔️ مواجهات اليوم {ordinal_day(day)}"]
    for p in data.get("pairs", []):
        mark = "🤝" if "تعادل" in p.get("status", "") else "✅"
        lines.append(f"{p['a']} {p['a_points']} - {p['b_points']} {p['b']} {mark} الفائز: {p['winner']}")
    if data.get("bye"):
        lines.append(f"🟡 راحة الجولة: {data['bye']}")
    return lines


def matchup_wins_map(start_day=1, end_day=31):
    wins = Counter()
    for day in get_existing_days(start_day, end_day):
        data = generate_matchups_for_day(day)
        for p in data.get("pairs", []):
            if p.get("winner"):
                wins[p["winner"]] += 1
        if data.get("bye"):
            # الراحة لا تحتسب فوزًا.
            pass
    return wins


def add_matchups_sheet_to_dashboard(xlsx_path, start_day=1, end_day=31):
    wb = load_workbook(xlsx_path)
    if "سجل المواجهات" in wb.sheetnames:
        del wb["سجل المواجهات"]
    ws = wb.create_sheet("سجل المواجهات")
    ws.sheet_view.rightToLeft = True
    ws.append(["اليوم", "المشارك 1", "نقاطه", "المشارك 2", "نقاطه", "النتيجة", "الفائز"])
    for day in get_existing_days(start_day, end_day):
        data = generate_matchups_for_day(day)
        if not data.get("pairs") and not data.get("bye"):
            ws.append([day, "لا توجد مواجهات", "", "لا يوجد مشاركون", "", "-", "-"])
            continue
        for p in data.get("pairs", []):
            ws.append([day, p["a"], p["a_points"], p["b"], p["b_points"], p.get("status", ""), p.get("winner", "")])
        if data.get("bye"):
            ws.append([day, data["bye"], "راحة", "-", "-", "راحة الجولة", data["bye"]])
    style_sheet(ws)
    wb.save(xlsx_path)


def build_daily_summary_text(day):
    data = build_day_summary(day)
    winners = " + ".join(data["winners"]) if data["winners"] else "لا يوجد"
    top_rows = data["sorted_rows"][:5]
    lines = [
        f"📊 ملخص اليوم {ordinal_day(day)}",
        f"👥 عدد المشاركين: {len(data['participants'])}",
        f"🏆 أسطورة اليوم: {winners} — {data['max_score']} نقطة",
        "",
        "🔥 أفضل 5:",
    ]
    for idx, r in enumerate(top_rows, start=1):
        lines.append(f"{idx}. {r['participant']} — {r['total']} نقطة")
    top_cap = data.get("top_captain")
    top_keep = data.get("top_keeper")
    if top_cap:
        lines.append(f"\n👑 أفضل كابتن: {top_cap['participant']} — +{top_cap['captain_points']}")
    if top_keep:
        lines.append(f"🧤 أفضل حارس: {top_keep['participant']} — +{top_keep['keeper_points']}")
    lines.append("")
    lines.extend(matchup_lines(day))
    return "\n".join(lines)


def create_daily_result_image(day, goals_count=None, clean_sheets=None):
    ensure_generated_dir()
    data = build_day_summary(day)
    rows = data["sorted_rows"]
    matchups = generate_matchups_for_day(day)
    extra_match_rows = len(matchups.get("pairs", [])) + (1 if matchups.get("bye") else 0)
    h = max(1180, 520 + len(rows) * 54 + extra_match_rows * 44 + 260)
    img, draw = make_canvas(1400, h)
    title_f = get_font(58)
    sub_f = get_font(34)
    small_f = get_font(25)
    row_f = get_font(27)

    draw_text(draw, (700, 85), f"فانتزي المصيف 2026 — اليوم {ordinal_day(day)}", title_f, fill="#FFFFFF")
    draw_text(draw, (700, 145), "نتائج اليوم وترتيب المشاركين", sub_f, fill="#FDE68A")

    winners = data["winners"]
    legends = " + ".join(winners) if winners else "لا يوجد"
    rounded_rect(draw, (80, 190, 1320, 340), radius=38, fill="#7C3AEDDD", outline="#FFFFFF33", width=2)
    draw_text(draw, (700, 235), "🏆 أسطورة اليوم 🏆", sub_f, fill="#FFFFFF")
    draw_text(draw, (700, 295), f"{legends} — {data['max_score']} نقطة", get_font(44), fill="#FFF6D6")

    top_cap = data["top_captain"]
    top_keep = data["top_keeper"]
    cards = [
        (80, 370, 455, 500, "👑 أفضل كابتن", f"{top_cap['participant']} +{top_cap['captain_points']}" if top_cap else "-", "#2563EB"),
        (512, 370, 887, 500, "🧤 أفضل حارس", f"{top_keep['participant']} +{top_keep['keeper_points']}" if top_keep else "-", "#10B981"),
        (945, 370, 1320, 500, "👥 المشاركون", f"{len(data['participants'])} مشارك", "#F59E0B"),
    ]
    for x1,y1,x2,y2,t,v,c in cards:
        rounded_rect(draw, (x1,y1,x2,y2), radius=28, fill=c+"DD", outline="#FFFFFF33", width=2)
        draw_text(draw, ((x1+x2)//2, y1+38), t, small_f, fill="#FFFFFF")
        draw_text(draw, ((x1+x2)//2, y1+92), v, sub_f, fill="#FFFFFF", max_width=x2-x1-20)

    y = 545
    rounded_rect(draw, (80, y, 1320, y+58), radius=18, fill="#FFFFFF22", outline="#FFFFFF30", width=1)
    headers = [(1180, "المشارك"), (850, "النقاط"), (640, "الحارس"), (410, "الكابتن"), (170, "المركز")]
    for x, t in headers:
        draw_text(draw, (x, y+30), t, small_f, fill="#FFFFFF")
    y += 70
    for idx, r in enumerate(rows, start=1):
        fill = "#FFFFFF16" if idx % 2 else "#FFFFFF0C"
        if idx == 1:
            fill = "#F2B70544"
        rounded_rect(draw, (80, y, 1320, y+48), radius=14, fill=fill, outline="#FFFFFF18", width=1)
        draw_text(draw, (1180, y+25), r["participant"], row_f, fill="#FFFFFF")
        draw_text(draw, (850, y+25), str(r["total"]), row_f, fill="#FDE68A")
        draw_text(draw, (640, y+25), f"+{r['keeper_points']}", row_f, fill="#A7F3D0")
        draw_text(draw, (410, y+25), f"+{r['captain_points']}", row_f, fill="#C4B5FD")
        medal = "🥇" if idx == 1 else "🥈" if idx == 2 else "🥉" if idx == 3 else str(idx)
        draw_text(draw, (170, y+25), medal, row_f, fill="#FFFFFF")
        y += 54

    y += 22
    rounded_rect(draw, (80, y, 1320, y+max(145, 58 + extra_match_rows*42)), radius=26, fill="#111827CC", outline="#FFFFFF22", width=2)
    draw_text(draw, (700, y+35), "⚔️ مواجهات اليوم", sub_f, fill="#FDE68A")
    yy = y + 82
    if not matchups.get("pairs") and not matchups.get("bye"):
        draw_text(draw, (700, yy), "لا توجد مواجهات — لا يوجد مشاركون", small_f, fill="#FFFFFF")
        yy += 42
    else:
        for p in matchups.get("pairs", []):
            line = f"{p['a']} {p['a_points']} - {p['b_points']} {p['b']}  |  الفائز: {p['winner']}"
            draw_text(draw, (700, yy), line, small_f, fill="#FFFFFF")
            yy += 42
        if matchups.get("bye"):
            draw_text(draw, (700, yy), f"🟡 راحة الجولة: {matchups['bye']}", small_f, fill="#FDE68A")
            yy += 42
    y = yy + 25

    box_h = 170
    rounded_rect(draw, (80, y, 670, y+box_h), radius=26, fill="#111827CC", outline="#FFFFFF22", width=2)
    rounded_rect(draw, (730, y, 1320, y+box_h), radius=26, fill="#111827CC", outline="#FFFFFF22", width=2)
    goals_lines = [f"{p} — {c}" for p,c in (goals_count or {}).items()] if goals_count else [f"{p} — {pts} نقطة" for p,pts in data["player_points"].most_common(5)] or ["لا يوجد"]
    clean_lines = clean_sheets or list(data["keeper_points"].keys()) or ["لا يوجد"]
    draw_text(draw, (375, y+35), "⚽ الهدافون", sub_f, fill="#FFFFFF")
    draw_text(draw, (375, y+105), "\n".join(goals_lines[:4]), small_f, fill="#E5E7EB", max_width=520)
    draw_text(draw, (1025, y+35), "🧤 الكلين شيت", sub_f, fill="#FFFFFF")
    draw_text(draw, (1025, y+105), "\n".join(clean_lines[:4]), small_f, fill="#E5E7EB", max_width=520)

    path = os.path.join(GENERATED_DIR, f"daily_result_day_{day}.png")
    img.save(path, quality=95)
    return path


def create_matches_image(day_name, matches):
    ensure_generated_dir()
    count = max(len(matches), 1)
    width = 1600
    row_h = 170 if count <= 3 else 140
    gap = 22 if count <= 3 else 16
    header_h = 230
    footer_h = 125
    height = max(900, header_h + count * row_h + max(0, count-1) * gap + footer_h + 40)
    img, draw = make_canvas(width, height)
    overlay = Image.new("RGBA", (width, height), (0,0,0,0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle((60,45,width-60,height-45), radius=48, fill="#02061770", outline="#FFFFFF18", width=2)
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)
    draw_text(draw, (width//2, 88), "مونديال المصيف 2026", get_font(58), fill="#FFFFFF")
    draw_text(draw, (width//2, 154), f"مباريات اليوم {day_name}", get_font(44), fill="#FDE68A")
    draw.line((210,205,width-210,205), fill="#FFFFFF35", width=2)
    colors = ["#7C3AED", "#2563EB", "#0891B2", "#059669", "#D97706", "#DC2626", "#4F46E5"]
    y = header_h
    for i, (a,b,t) in enumerate(matches, start=1):
        c = colors[(i-1)%len(colors)]
        rounded_rect(draw, (90,y,width-90,y+row_h), radius=36, fill=c+"D5", outline="#FFFFFF30", width=2)
        flag_w = 180 if count <= 3 else 145
        flag_h = 110 if count <= 3 else 90
        cy = y + row_h//2
        paste_flag(img, a, (width-285, cy-flag_h//2, width-105, cy+flag_h//2))
        paste_flag(img, b, (105, cy-flag_h//2, 285, cy+flag_h//2))
        name_font = get_font(46 if count <= 3 else 38)
        draw_text(draw, (width-470, cy), a, name_font, fill="#FFFFFF", max_width=360)
        draw_text(draw, (470, cy), b, name_font, fill="#FFFFFF", max_width=360)
        draw_text(draw, (width//2, cy-20), "×", get_font(58 if count <= 3 else 48), fill="#FDE68A")
        if t:
            badge_w = 210 if count <= 3 else 180
            badge_h = 44 if count <= 3 else 38
            rounded_rect(draw, (width//2-badge_w//2, cy+32, width//2+badge_w//2, cy+32+badge_h), radius=18, fill="#020617E6", outline="#FFFFFF40", width=1)
            draw_text(draw, (width//2, cy+32+badge_h//2), t, get_font(26 if count <= 3 else 22), fill="#FFFFFF")
        y += row_h + gap
    footer_y = min(height-86, y+55)
    draw_text(draw, (width//2, footer_y), "المصيف ينقل لكم الحدث", get_font(42), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, f"matches_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path


def parse_match_results_text(text):
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    results = []
    for line in lines[1:]:
        if "|" in line:
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 4:
                try:
                    results.append((parts[0], int(parts[1]), int(parts[2]), parts[3]))
                    continue
                except Exception:
                    pass
            if len(parts) == 2:
                m1 = re.match(r"(.+?)(\d+)\s*$", parts[0])
                m2 = re.match(r"^(\d+)\s*(.+)$", parts[1])
                if m1 and m2:
                    results.append((normalize_name(m1.group(1)), int(m1.group(2)), int(m2.group(1)), normalize_name(m2.group(2))))
        m = re.match(r"(.+?)\s+(\d+)\s*[-–:]\s*(\d+)\s+(.+)$", line)
        if m:
            results.append((normalize_name(m.group(1)), int(m.group(2)), int(m.group(3)), normalize_name(m.group(4))))
    return results


def create_match_results_image(results):
    ensure_generated_dir()
    width = 1600
    row_h = 165 if len(results) <= 3 else 135
    height = max(820, 230 + len(results)*row_h + 180)
    img, draw = make_canvas(width, height)
    draw_text(draw, (width//2, 92), "مونديال المصيف 2026", get_font(64), fill="#FFFFFF")
    draw_text(draw, (width//2, 160), "نتائج مباريات اليوم", get_font(46), fill="#FDE68A")
    y = 240
    colors = ["#7C3AED", "#2563EB", "#0891B2", "#059669", "#D97706"]
    for i,(a,sa,sb,b) in enumerate(results, start=1):
        c = colors[(i-1)%len(colors)]
        rounded_rect(draw, (90,y,width-90,y+row_h), radius=36, fill=c+"D5", outline="#FFFFFF30", width=2)
        cy = y + row_h//2
        paste_flag(img, a, (width-280, cy-55, width-105, cy+55))
        paste_flag(img, b, (105, cy-55, 280, cy+55))
        draw_text(draw, (width-480, cy), a, get_font(44), fill="#FFFFFF", max_width=340)
        draw_text(draw, (480, cy), b, get_font(44), fill="#FFFFFF", max_width=340)
        rounded_rect(draw, (width//2-125, cy-44, width//2+125, cy+44), radius=26, fill="#020617E6", outline="#FDE68A", width=2)
        draw_text(draw, (width//2, cy), f"{sa} - {sb}", get_font(54), fill="#FDE68A")
        y += row_h + 24
    draw_text(draw, (width//2, height-70), "المصيف ينقل لكم الحدث", get_font(42), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, "match_results_today.png")
    img.save(path, quality=95)
    return path


def build_match_results_caption(results):
    lines = ["🏆 نتائج مباريات اليوم 🏆", ""]
    for a,sa,sb,b in results:
        lines.append(f"{a} {sa} - {sb} {b}")
    lines.append("\nالمصيف ينقل لكم الحدث")
    return "\n".join(lines)


async def match_results_today_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    results = parse_match_results_text(update.message.text)
    if not results:
        await update.message.reply_text("اكتبها كذا:\n/نتائج_مباريات_اليوم\nالسعودية|2|1|إسبانيا\nفرنسا|3|0|البرازيل")
        return
    try:
        path = create_match_results_image(results)
        await send_photo_path(update, path, build_match_results_caption(results))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم نتائج المباريات ❌\n{e}")


def parse_scorers_text(text):
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    items = []
    for line in lines[1:]:
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 2:
            name = parts[0]
            goals = score_to_int(parts[1])
            team = parts[2] if len(parts) >= 3 else ""
            if name and goals > 0:
                items.append((name, goals, team))
    return sorted(items, key=lambda x: (-x[1], x[0]))


def create_top_scorers_image(items):
    ensure_generated_dir()
    width = 1400
    row_h = 110
    height = max(850, 230 + len(items[:10])*row_h + 150)
    img, draw = make_canvas(width, height)
    draw_text(draw, (width//2, 85), "مونديال المصيف 2026", get_font(60), fill="#FFFFFF")
    draw_text(draw, (width//2, 150), "هدافين البطولة حتى الآن", get_font(44), fill="#FDE68A")
    y = 235
    for i,(name,goals,team) in enumerate(items[:10], start=1):
        fill = "#F2B70555" if i == 1 else "#FFFFFF16"
        rounded_rect(draw, (90,y,width-90,y+88), radius=28, fill=fill, outline="#FFFFFF25", width=2)
        draw_text(draw, (1220, y+44), str(i), get_font(46), fill="#FDE68A" if i == 1 else "#FFFFFF")
        if team:
            paste_flag(img, team, (1040, y+12, 1130, y+76))
        draw_text(draw, (790, y+44), name, get_font(42), fill="#FFFFFF", max_width=430)
        draw_text(draw, (250, y+44), f"{goals} {'هدف' if goals == 1 else 'أهداف'}", get_font(38), fill="#FDE68A")
        y += row_h
    draw_text(draw, (width//2, height-70), "المصيف ينقل لكم الحدث", get_font(38), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, "top_scorers.png")
    img.save(path, quality=95)
    return path


def build_top_scorers_caption(items):
    lines = ["⚽ هدافين البطولة حتى الآن", ""]
    for i,(name,goals,team) in enumerate(items, start=1):
        team_txt = f" — {team}" if team else ""
        lines.append(f"{i}. {name}{team_txt} — {goals} {'هدف' if goals == 1 else 'أهداف'}")
    lines.append("\nالمصيف ينقل لكم الحدث")
    return "\n".join(lines)


async def top_scorers_design_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    items = parse_scorers_text(update.message.text)
    if not items:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_هدافين\nراؤول خيمينيز|4|المكسيك\nفلوريان فيرتز|3|ألمانيا")
        return
    try:
        path = create_top_scorers_image(items)
        await send_photo_path(update, path, build_top_scorers_caption(items))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم الهدافين ❌\n{e}")


def find_participant_name(query):
    q = normalize_name(query)
    if not q:
        return None
    for name in PARTICIPANTS:
        if normalize_name(name) == q:
            return name
    for name in PARTICIPANTS:
        if q in normalize_name(name) or normalize_name(name) in q:
            return name
    return None


def participant_best_captain(name, start_day=1, end_day=31):
    best = ("-", 0, "-")
    for day in get_existing_days(start_day, end_day):
        for r in read_day_rows(day):
            if r["participant"] == name and r.get("captain_points", 0) > best[1]:
                best = (r.get("captain", "-"), r.get("captain_points", 0), day)
    return best


def create_participant_card_image(name, start_day=1, end_day=31):
    ensure_generated_dir()
    stats = collect_stats(start_day, end_day)
    if name not in PARTICIPANTS:
        raise ValueError("اسم المشارك غير موجود")
    rank = stats["ranking"].index(name) + 1 if name in stats["ranking"] else "-"
    total = stats["totals"].get(name, 0)
    day_scores = stats["scores_by_day"].get(name, {})
    best_day = max(day_scores, key=lambda d: day_scores[d]) if day_scores else "-"
    worst_day = min(day_scores, key=lambda d: day_scores[d]) if day_scores else "-"
    best_score = day_scores.get(best_day, 0) if best_day != "-" else 0
    worst_score = day_scores.get(worst_day, 0) if worst_day != "-" else 0
    best_cap, best_cap_pts, best_cap_day = participant_best_captain(name, start_day, end_day)
    wins = matchup_wins_map(start_day, end_day).get(name, 0)
    pc = stats["participation_count"].get(name, 0)
    days_count = len(stats["days"])
    pct = f"{round(pc / days_count * 100, 1)}%" if days_count else "0%"

    img, draw = make_canvas(1200, 900)
    draw_text(draw, (600, 85), "بطاقة مشارك فانتزي المصيف", get_font(52), fill="#FFFFFF")
    rounded_rect(draw, (90,145,1110,270), radius=38, fill="#7C3AEDDD", outline="#FFFFFF33", width=2)
    draw_text(draw, (600, 205), name, get_font(64), fill="#FDE68A")
    cards = [
        (90,310,360,430,"المركز", f"#{rank}"),
        (465,310,735,430,"النقاط", str(total)),
        (840,310,1110,430,"أسطورة اليوم", str(stats["daily_wins"].get(name,0))),
        (90,470,360,590,"أفضل يوم", f"{best_day} — {best_score}"),
        (465,470,735,590,"أسوأ يوم", f"{worst_day} — {worst_score}"),
        (840,470,1110,590,"نسبة المشاركة", pct),
        (90,630,545,750,"أفضل كابتن", f"{best_cap} +{best_cap_pts}"),
        (655,630,1110,750,"فوز المواجهات", str(wins)),
    ]
    for x1,y1,x2,y2,t,v in cards:
        rounded_rect(draw,(x1,y1,x2,y2), radius=28, fill="#111827CC", outline="#FFFFFF25", width=2)
        draw_text(draw, ((x1+x2)//2, y1+34), t, get_font(26), fill="#E5E7EB")
        draw_text(draw, ((x1+x2)//2, y1+82), v, get_font(36), fill="#FFFFFF", max_width=x2-x1-24)
    draw_text(draw, (600, 820), "فانتزي المصيف 2026", get_font(34), fill="#FDE68A")
    path = os.path.join(GENERATED_DIR, f"participant_card_{_safe_filename(name)}.png")
    img.save(path, quality=95)
    return path


async def participant_card_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text or ""
    query = re.sub(r"^/بطاقة\s*", "", text).strip()
    name = find_participant_name(query)
    if not name:
        await update.message.reply_text("اكتب اسم المشارك بعد الأمر، مثال:\n/بطاقة فارس سالم")
        return
    try:
        path = create_participant_card_image(name)
        await send_photo_path(update, path, f"بطاقة {name} ✅")
    except Exception as e:
        await update.message.reply_text(f"تعذر إنشاء البطاقة ❌\n{e}")


def period_stats(start_day, end_day):
    days = get_existing_days(start_day, end_day)
    totals = Counter()
    cap_points = Counter()
    keeper_points = Counter()
    player_impact = Counter()
    player_zero_popularity = Counter()
    active_counts = Counter()
    for day in days:
        for r in read_day_rows(day):
            if r.get("participated"):
                totals[r["participant"]] += r["total"]
                active_counts[r["participant"]] += 1
                cap_points[r["participant"]] += r.get("captain_points", 0)
                if not is_no_participation(r.get("keeper")):
                    keeper_points[r["keeper"]] += r.get("keeper_points", 0)
                for pk, ptsk in (("p1","p1_points"),("p2","p2_points"),("p3","p3_points")):
                    p = r.get(pk)
                    if not is_no_participation(p):
                        pts = r.get(ptsk,0)
                        extra = r.get("captain_points",0) if r.get("captain") == p else 0
                        player_impact[p] += pts + extra
                        if pts == 0:
                            player_zero_popularity[p] += 1
    return days, totals, cap_points, keeper_points, player_impact, player_zero_popularity, active_counts


def resolve_cup_match(a, b, day, ranks_before):
    rows = {r["participant"]: r for r in read_day_rows(day)}
    ra, rb = rows.get(a), rows.get(b)
    a_play = bool(ra and ra.get("participated"))
    b_play = bool(rb and rb.get("participated"))
    pa = ra.get("total", 0) if ra else 0
    pb = rb.get("total", 0) if rb else 0
    if a_play and not b_play:
        return a, pa, pb, "انسحاب الخصم"
    if b_play and not a_play:
        return b, pa, pb, "انسحاب الخصم"
    if not a_play and not b_play:
        winner = a if ranks_before.get(a, 999) <= ranks_before.get(b, 999) else b
        return winner, pa, pb, "الاثنان لم يشاركا — حُسمت بالأفضلية"
    if pa > pb:
        return a, pa, pb, "فوز بالنقاط"
    if pb > pa:
        return b, pa, pb, "فوز بالنقاط"
    winner = a if ranks_before.get(a, 999) <= ranks_before.get(b, 999) else b
    return winner, pa, pb, "تعادل — حُسمت بالأفضلية"


def compute_fantasy_cup(start_day, end_day):
    """كأس كل 4 أيام: دور تمهيدي ثم ربع ثم نصف ثم نهائي. مناسب لـ12 مشارك."""
    ranks = previous_ranking_before_day(start_day)
    seeds = sorted(PARTICIPANTS, key=lambda n: ranks.get(n, 999))
    days = list(range(int(start_day), int(end_day)+1))
    if len(days) < 4:
        return {"champion": "-", "rounds": [], "note": "الفترة أقل من 4 أيام"}
    d1,d2,d3,d4 = days[:4]
    rounds = []
    # دور 12: أول 4 راحة، 5-12 يلعبون
    bye4 = seeds[:4]
    prelim_pairs = [(seeds[4], seeds[11]), (seeds[5], seeds[10]), (seeds[6], seeds[9]), (seeds[7], seeds[8])] if len(seeds) >= 12 else []
    prelim_winners = []
    for a,b in prelim_pairs:
        w,pa,pb,st = resolve_cup_match(a,b,d1,ranks)
        prelim_winners.append(w)
        rounds.append({"round":"دور 12", "day":d1, "a":a, "b":b, "pa":pa, "pb":pb, "winner":w, "status":st})
    q_players = [bye4[0], prelim_winners[0], bye4[3], prelim_winners[3], bye4[1], prelim_winners[1], bye4[2], prelim_winners[2]] if len(prelim_winners)==4 else seeds[:8]
    q_pairs = [(q_players[0],q_players[1]), (q_players[2],q_players[3]), (q_players[4],q_players[5]), (q_players[6],q_players[7])]
    q_winners = []
    for a,b in q_pairs:
        w,pa,pb,st = resolve_cup_match(a,b,d2,ranks)
        q_winners.append(w)
        rounds.append({"round":"ربع النهائي", "day":d2, "a":a, "b":b, "pa":pa, "pb":pb, "winner":w, "status":st})
    s_pairs = [(q_winners[0], q_winners[1]), (q_winners[2], q_winners[3])]
    s_winners = []
    for a,b in s_pairs:
        w,pa,pb,st = resolve_cup_match(a,b,d3,ranks)
        s_winners.append(w)
        rounds.append({"round":"نصف النهائي", "day":d3, "a":a, "b":b, "pa":pa, "pb":pb, "winner":w, "status":st})
    a,b = s_winners[0], s_winners[1]
    w,pa,pb,st = resolve_cup_match(a,b,d4,ranks)
    rounds.append({"round":"النهائي", "day":d4, "a":a, "b":b, "pa":pa, "pb":pb, "winner":w, "status":st})
    return {"champion": w, "rounds": rounds, "note": ""}


def create_period_report_image(start_day, end_day):
    ensure_generated_dir()
    days, totals, cap_points, keeper_points, player_impact, player_zero, active_counts = period_stats(start_day, end_day)
    if not days:
        raise ValueError("ما فيه أيام في الفترة")
    champion = totals.most_common(1)[0][0] if totals else "-"
    best_participant = champion
    best_cap = cap_points.most_common(1)[0] if cap_points else ("-",0)
    best_keeper = keeper_points.most_common(1)[0] if keeper_points else ("-",0)
    best_player = player_impact.most_common(1)[0] if player_impact else ("-",0)
    disappointment = player_zero.most_common(1)[0] if player_zero else ("-",0)
    mw = matchup_wins_map(start_day, end_day).most_common(1)
    king_matchups = mw[0] if mw else ("-",0)
    cup = compute_fantasy_cup(start_day, end_day)
    cup_champ = cup.get("champion", "-")

    img, draw = make_canvas(1400, 1050)
    draw_text(draw, (700,80), f"تقرير الفترة من اليوم {start_day} إلى {end_day}", get_font(54), fill="#FFFFFF")
    draw_text(draw, (700,140), "فانتزي المصيف 2026", get_font(38), fill="#FDE68A")
    cards = [
        (80,200,430,330,"🏆 بطل الفترة", f"{champion}\n{totals.get(champion,0)} نقطة"),
        (525,200,875,330,"👑 أفضل كابتن", f"{best_cap[0]}\n+{best_cap[1]}"),
        (970,200,1320,330,"🧤 أفضل حارس", f"{best_keeper[0]}\n{best_keeper[1]} نقطة"),
        (80,380,430,510,"🔥 أكثر لاعب أفاد", f"{best_player[0]}\n{best_player[1]} نقطة"),
        (525,380,875,510,"😅 خيبة الفترة", f"{disappointment[0]}\n{disappointment[1]} اختيارات صفر"),
        (970,380,1320,510,"⚔️ ملك المواجهات", f"{king_matchups[0]}\n{king_matchups[1]} فوز"),
        (80,560,1320,690,"🏆 بطل كأس الفانتزي", cup_champ),
    ]
    for x1,y1,x2,y2,t,v in cards:
        rounded_rect(draw,(x1,y1,x2,y2), radius=30, fill="#111827CC", outline="#FFFFFF25", width=2)
        draw_text(draw, ((x1+x2)//2, y1+38), t, get_font(27), fill="#E5E7EB")
        draw_text(draw, ((x1+x2)//2, y1+88), v, get_font(38), fill="#FFFFFF", max_width=x2-x1-30)

    y = 740
    draw_text(draw, (700,y), "أفضل 5 في الفترة", get_font(36), fill="#FDE68A")
    y += 55
    for i,(name,pts) in enumerate(totals.most_common(5), start=1):
        rounded_rect(draw,(170,y,1230,y+55), radius=18, fill="#FFFFFF16", outline="#FFFFFF20", width=1)
        draw_text(draw,(1110,y+28), f"{i}. {name}", get_font(30), fill="#FFFFFF")
        draw_text(draw,(300,y+28), f"{pts} نقطة", get_font(30), fill="#FDE68A")
        y += 65
    path = os.path.join(GENERATED_DIR, f"period_report_{start_day}_{end_day}.png")
    img.save(path, quality=95)
    return path


async def period_report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nums = get_numbers(update.message.text)
    if len(nums) >= 2:
        start_day, end_day = nums[0], nums[1]
    else:
        await update.message.reply_text("اكتب الفترة، مثال:\n/تقرير_الفترة 1 4")
        return
    if start_day > end_day:
        start_day, end_day = end_day, start_day
    try:
        path = create_period_report_image(start_day, end_day)
        await send_photo_path(update, path, f"تقرير الفترة {start_day} - {end_day} ✅")
    except Exception as e:
        await update.message.reply_text(f"تعذر إنشاء تقرير الفترة ❌\n{e}")


async def dashboard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        nums = get_numbers(update.message.text)
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
            await update.message.reply_text("ما لقيت أيام لإحصائيات الداشبورد.")
            return
        add_matchups_sheet_to_dashboard(file_name, start_day, end_day)
        caption = (
            "تم إنشاء ملف الإحصائيات الكامل ✅\n"
            f"النطاق: من اليوم {start_day} إلى اليوم {end_day}\n"
            f"الأيام المحسوبة: {', '.join(map(str, stats['days']))}\n\n"
            "الصفحات تشمل سجل المواجهات ✅"
        )
        with open(file_name, "rb") as file:
            await update.message.reply_document(document=file, filename=file_name, caption=caption)
    except Exception as e:
        await update.message.reply_text(f"صار خطأ أثناء إنشاء الإحصائيات ❌\n\nالسبب:\n{e}")


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
    backup_files(f"before_results_day_{day}", files=[excel_file(day), LOCKED_FILE, MATCHUPS_FILE])
    file_name = calculate_points(day, goals_count, clean_sheets)
    # المواجهات بعد النتائج فقط
    generate_matchups_for_day(day, force=True)
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
        f"🏆 أسطورة اليوم: {legends_text} — {max_score} نقطة\n\n"
        + "\n".join(matchup_lines(day))
    )
    if warnings:
        caption += "\n\n" + "\n\n".join(warnings)
    with open(file_name, "rb") as file:
        await update.message.reply_document(document=file, filename=file_name, caption=caption)
    await update.message.reply_text(build_result_announcement(day, winners))
    if load_settings().get("auto_images", True):
        try:
            path = create_daily_result_image(day, goals_count=goals_count, clean_sheets=clean_sheets)
            await send_photo_path(update, path, build_result_announcement(day, winners))
        except Exception as e:
            await update.message.reply_text(f"تم حساب النتائج، لكن تعذر إنشاء الصورة التلقائية ❌\n{e}")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "الأوامر الأساسية:\n"
        "/اضافه 5\n/نتائج 5\n/احصائيات\n/احصائيات 1 6\n/ترتيب_نص\n\n"
        "أوامر الصور والتقارير:\n"
        "/صورة_اليوم 6\n/صورة_الترتيب 1 6\n/صورة_الاساطير 1 6\n/صورة_احصائيات 1 6 لوحة عامة\n/صور_الاحصائيات 1 6\n"
        "/بطاقة فارس سالم\n/تقرير_الفترة 1 4\n/تفعيل_الصور_التلقائية\n/إيقاف_الصور_التلقائية\n\n"
        "أوامر التصاميم:\n"
        "/تصميم_مباريات\n/تصميم_مباريات_تلقائي\n"
        "/تصميم_نتائج_مباريات\n/تصميم_نتائج_مباريات_تلقائي\n"
        "/تصميم_ترتيب_مجموعة\n/تصميم_ترتيب_مجموعة_تلقائي\n"
        "/تصميم_هدافين\n/تصميم_هدافين_تلقائي\n"
        "\nالتصميم الإضافي V24:\n"
        "/تصميم_مباريات_ستايل2\n/تصميم_نتائج_مباريات_ستايل2\n"
        "/تصميم_هدافين_ستايل2\n/تصميم_ترتيب_مجموعة_ستايل2\n"
        "/تصميم_مباريات_اطار\n/تصميم_نتائج_مباريات_اطار\n"
        "/تصميم_جميع_المجموعات\n\n"
        "أوامر الفحص والنشر:\n"
        "/الأيام\n/فحص 5\n/مشاركين 5\n/اسطورة 5\n/مقارنة 4 5\n/اعلان_اليوم 5\n/ملخص_اليوم 5\n\n"
        "أوامر الاستيراد والنسخ:\n"
        "/استيراد_ملف — أرسل ملف الإكسل لحاله ثم اكتب الأمر\n/اعتماد_استيراد\n/إلغاء_استيراد\n/نسخة_احتياطية\n/استرجاع_نسخة — أرسل ملف ZIP لحاله ثم اكتب الأمر\n/تنظيف_الأيام\n/تنظيف_الملفات\n\n"
        "أوامر الأمان:\n"
        "/مسح_نتائج 5\n/مسح_يوم 5\n/مسح_الكل تأكيد\n/استرجاع_آخر\n/قفل_يوم 5\n/فتح_يوم 5\n/معرفي"
    )




# ============================================================
# V13 — تصاميم القوالب + التصاميم التلقائية
# الأوامر:
# /تصميم_مباريات + /تصميم_مباريات_تلقائي
# /تصميم_نتائج_مباريات + /تصميم_نتائج_مباريات_تلقائي
# /تصميم_ترتيب_مجموعة + /تصميم_ترتيب_مجموعة_تلقائي
# /تصميم_هدافين + /تصميم_هدافين_تلقائي
# ============================================================

DESIGN_ZIP = "design_assets.zip"
TEMPLATE_DIR = os.path.join("assets", "templates")

def ensure_design_assets():
    logo_ok = os.path.exists(os.path.join("assets", "logos", "masif_logo_mark.png"))
    templates_ok = os.path.exists(TEMPLATE_DIR)
    if templates_ok and logo_ok:
        return
    if os.path.exists(DESIGN_ZIP):
        try:
            import zipfile
            with zipfile.ZipFile(DESIGN_ZIP, "r") as z:
                z.extractall(".")
        except Exception:
            pass

def template_path(name):
    ensure_design_assets()
    p = os.path.join(TEMPLATE_DIR, name)
    return p if os.path.exists(p) else None

def design_canvas(template_name=None, width=1200, height=1500, theme="purple"):
    """
    V16 broadcast-style canvas:
    - very dark background
    - soft green/red/purple edge glows
    - thin inner colored frame
    """
    if template_name:
        p = template_path(template_name)
        if p and os.path.exists(p):
            try:
                img = Image.open(p).convert("RGB").resize((width, height), Image.LANCZOS)
                return img, ImageDraw.Draw(img)
            except Exception:
                pass

    img = Image.new("RGB", (width, height), "#05070D")
    draw = ImageDraw.Draw(img)
    for y in range(height):
        t = y / max(height, 1)
        val = int(5 + 11 * t)
        draw.line((0, y, width, y), fill=(val, val + 2, val + 8))

    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    theme_colors = {
        "purple": ("#22C55E", "#EF4444", "#7C3AED"),
        "blue": ("#06B6D4", "#22C55E", "#EF4444"),
        "gold": ("#F59E0B", "#7C3AED", "#22C55E"),
    }
    left, right, bottom = theme_colors.get(theme, theme_colors["purple"])

    def hx(c):
        c = c.lstrip("#")
        return tuple(int(c[i:i+2], 16) for i in (0, 2, 4))

    od.ellipse((-250, 65, 285, 610), fill=(*hx(left), 70))
    od.ellipse((width-285, 85, width+250, 640), fill=(*hx(right), 65))
    od.ellipse((-170, height-500, 380, height+70), fill=(*hx(bottom), 42))
    od.ellipse((width-380, height-500, width+170, height+70), fill=(*hx("#0EA5E9"), 35))
    overlay = overlay.filter(ImageFilter.GaussianBlur(88))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    return img, ImageDraw.Draw(img)


def event_logo_path():
    ensure_design_assets()
    for p in [
        os.path.join("assets", "logos", "masif_logo_mark.png"),
        os.path.join("assets", "logos", "masif_logo_full.png"),
    ]:
        if os.path.exists(p):
            return p
    return None

def paste_event_logo(img, width, y=48):
    """
    شعار استراحة المصيف بجانب العنوان.
    يستخدم نسخة mark لأنها أوضح بحجم صغير.
    """
    p = event_logo_path()
    if not p:
        return
    try:
        logo = Image.open(p).convert("RGBA")
        logo.thumbnail((130, 100), Image.LANCZOS)
        badge_w = max(108, logo.width + 24)
        badge_h = max(88, logo.height + 18)
        x = width - badge_w - 64

        # Badge خلف الشعار عشان يبان فوق الخلفية الداكنة
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        od.rounded_rectangle((x, y, x + badge_w, y + badge_h), radius=24, fill=(5, 7, 13, 135), outline=(253, 230, 138, 90), width=1)
        img_rgba = img.convert("RGBA")
        img_rgba = Image.alpha_composite(img_rgba, overlay)
        img.paste(img_rgba.convert("RGB"))

        lx = x + (badge_w - logo.width) // 2
        ly = y + (badge_h - logo.height) // 2
        img.paste(logo, (lx, ly), logo)
    except Exception:
        pass

def clean_group_title_for_design(title):
    title = str(title or "").strip()
    # منع ظهور مربع بدل حرف المجموعة في بعض الخطوط/الأجهزة
    mapping = {
        "A": "أ", "B": "ب", "C": "ج", "D": "د", "E": "هـ", "F": "و",
        "G": "ز", "H": "ح", "I": "ط", "J": "ي", "K": "ك", "L": "ل",
    }
    def repl(m):
        letter = m.group(1).upper()
        return "المجموعة " + mapping.get(letter, letter)
    title = re.sub(r"المجموعة\s+([A-Za-z])\b", repl, title)
    title = re.sub(r"Group\s+([A-Za-z])\b", repl, title, flags=re.I)
    return title

def draw_design_header(draw, width, title, subtitle, img=None):
    # V23: بدون شعار في جميع التصاميم
    draw_text(draw, (width//2, 96), title, get_font(54), fill="#FFFFFF")
    draw_text(draw, (width//2, 160), subtitle, get_font(34), fill="#FDE68A")
    draw.line((210, 213, width-210, 213), fill="#FFFFFF45", width=1)

def footer_event(draw, width, height):
    draw.line((300, height-92, width-300, height-92), fill="#FFFFFF35", width=1)
    draw_text(draw, (width//2, height-58), "المصيف ينقل لكم الحدث", get_font(30), fill="#FFFFFF")


def draw_broadcast_inner_frame(draw, width, height, top=250, bottom_pad=118, accent="#22C55E"):
    # إطار خفيف جدًا جدًا — المحتوى كله داخله
    left = 50
    right = width - 50
    bottom = height - bottom_pad
    try:
        draw.rounded_rectangle((left, top, right, bottom), radius=38, outline="#FFFFFF55", width=1)
        draw.rounded_rectangle((left+3, top+3, right-3, bottom-3), radius=35, outline=accent, width=1)
    except Exception:
        draw.rectangle((left, top, right, bottom), outline="#FFFFFF55", width=1)
    return (left, top, right, bottom)

def v16_row_style(index):
    # ألوان صفوف رياضية واضحة داخل الإطار
    return ["#151A26", "#0B1020", "#111827", "#101820"][index % 4]

def v16_accent(index):
    return ["#22C55E", "#EF4444", "#7C3AED", "#06B6D4", "#F59E0B", "#2563EB"][index % 6]

def v16_fit_row_metrics(count, kind="match"):
    # يعطي راحته إذا مباراة وحدة أو 10، ويصغّر تلقائيًا عند الكثرة
    if kind == "standings":
        if count <= 4:
            return 120, 18, 40
        if count <= 6:
            return 108, 14, 36
        return 92, 10, 30
    if kind == "scorers":
        if count <= 5:
            return 118, 18, 38
        if count <= 8:
            return 102, 13, 32
        return 88, 10, 28
    # match/results
    if count <= 2:
        return 145, 24, 40
    if count <= 5:
        return 126, 18, 34
    if count <= 8:
        return 104, 13, 29
    return 90, 10, 25

def parse_match_results_design_text(text):
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    if len(lines) <= 1:
        return "اليوم", []
    idx = 1
    day_name = "اليوم"
    # إذا السطر الثاني ليس نتيجة، نعتبره اسم اليوم
    if idx < len(lines) and "|" not in lines[idx] and not re.search(r"\d+\s*[-–:]\s*\d+", lines[idx]):
        day_name = lines[idx]
        idx += 1
    results = []
    for line in lines[idx:]:
        if "|" in line:
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 4:
                try:
                    results.append((parts[0], int(parts[1]), int(parts[2]), parts[3]))
                    continue
                except Exception:
                    pass
            if len(parts) == 2:
                m1 = re.match(r"(.+?)(\d+)\s*$", parts[0])
                m2 = re.match(r"^(\d+)\s*(.+)$", parts[1])
                if m1 and m2:
                    results.append((normalize_name(m1.group(1)), int(m1.group(2)), int(m2.group(1)), normalize_name(m2.group(2))))
                    continue
        m = re.match(r"(.+?)\s+(\d+)\s*[-–:]\s*(\d+)\s+(.+)$", line)
        if m:
            results.append((normalize_name(m.group(1)), int(m.group(2)), int(m.group(3)), normalize_name(m.group(4))))
    return day_name, results

def parse_group_standing_text(text):
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    if len(lines) <= 1:
        return "المجموعة", []
    group_title = lines[1]
    rows = []
    for line in lines[2:]:
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 4:
            team = parts[0]
            played = cell_int(parts[1], 0)
            diff = cell_int(parts[2], 0)
            pts = cell_int(parts[3], 0)
            rows.append((team, played, diff, pts))
    rows.sort(key=lambda x: (x[3], x[2], x[0]), reverse=True)
    return group_title, rows

def create_matches_template_image(day_name, matches, use_template=True):
    ensure_generated_dir()
    count = max(len(matches), 1)
    width = 1200
    row_h, gap, name_size = v16_fit_row_metrics(count, "match")
    # تكبير بسيط للصفوف إذا العدد قليل
    if count <= 2:
        row_h += 12
        name_size += 2
    frame_top = 245
    content_h = count * row_h + max(0, count - 1) * gap
    height = max(760, frame_top + content_h + 195)
    img, draw = design_canvas("matches_template.png" if use_template else None, width, height, "purple")
    draw_design_header(draw, width, "مونديال المصيف 2026", f"مباريات اليوم {day_name}", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=112, accent="#22C55E")

    available_h = (fy2 - fy1) - 70
    y = fy1 + 38 + max(0, (available_h - content_h) // 2)
    for i, (a, b, t) in enumerate(matches, start=1):
        accent = v16_accent(i-1)
        rounded_rect(draw, (92, y, width-92, y+row_h), radius=28, fill="#0B1020", outline=accent, width=2)
        cy = y + row_h//2

        # V17: تكبير الأعلام
        flag_w = min(150, max(94, row_h - 18))
        paste_flag(img, a, (width-238, cy-flag_w//2, width-238+flag_w, cy+flag_w//2))
        paste_flag(img, b, (238-flag_w, cy-flag_w//2, 238, cy+flag_w//2))

        draw_text(draw, (width-400, cy-4), a, get_font(name_size), fill="#FFFFFF", max_width=285)
        draw_text(draw, (400, cy-4), b, get_font(name_size), fill="#FFFFFF", max_width=285)

        draw_text(draw, (width//2, cy-20), "×", get_font(max(40, name_size+18)), fill="#FDE68A")
        if t:
            badge_w = 214 if row_h >= 104 else 178
            badge_h = 46 if row_h >= 104 else 36
            rounded_rect(draw, (width//2-badge_w//2, cy+24, width//2+badge_w//2, cy+24+badge_h), radius=16, fill="#05070D", outline="#FFFFFF99", width=1)
            draw_text(draw, (width//2, cy+24+badge_h//2), t, get_font(max(20, name_size-12)), fill="#FFFFFF")

        y += row_h + gap

    footer_event(draw, width, height)
    suffix = "template" if use_template else "auto"
    path = os.path.join(GENERATED_DIR, f"matches_{suffix}_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path

def create_match_results_template_image(day_name, results, use_template=True):
    ensure_generated_dir()
    count = max(len(results), 1)
    width = 1200
    row_h, gap, name_size = v16_fit_row_metrics(count, "match")
    if count <= 2:
        row_h += 10
        name_size += 2
    frame_top = 245
    content_h = count * row_h + max(0, count - 1) * gap
    height = max(760, frame_top + content_h + 195)
    img, draw = design_canvas("match_results_template.png" if use_template else None, width, height, "blue")
    draw_design_header(draw, width, "مونديال المصيف 2026", f"نتائج مباريات اليوم {day_name}", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=112, accent="#06B6D4")

    available_h = (fy2 - fy1) - 70
    y = fy1 + 38 + max(0, (available_h - content_h) // 2)
    for i, (a, sa, sb, b) in enumerate(results, start=1):
        accent = v16_accent(i)
        rounded_rect(draw, (92, y, width-92, y+row_h), radius=28, fill="#0B1020", outline=accent, width=2)
        cy = y + row_h//2

        flag_w = min(146, max(92, row_h - 18))
        paste_flag(img, a, (width-232, cy-flag_w//2, width-232+flag_w, cy+flag_w//2))
        paste_flag(img, b, (232-flag_w, cy-flag_w//2, 232, cy+flag_w//2))

        draw_text(draw, (width-395, cy), a, get_font(name_size), fill="#FFFFFF", max_width=280)
        draw_text(draw, (395, cy), b, get_font(name_size), fill="#FFFFFF", max_width=280)

        badge_w = 230 if row_h >= 104 else 190
        badge_h = 70 if row_h >= 104 else 56
        rounded_rect(draw, (width//2-badge_w//2, cy-badge_h//2, width//2+badge_w//2, cy+badge_h//2), radius=22, fill="#05070D", outline="#FDE68A", width=2)
        draw_text(draw, (width//2, cy), f"{sb} - {sa}", get_font(max(32, name_size+12)), fill="#FDE68A")
        y += row_h + gap

    footer_event(draw, width, height)
    suffix = "template" if use_template else "auto"
    path = os.path.join(GENERATED_DIR, f"match_results_{suffix}_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path

def create_group_standing_image(group_title, rows, use_template=True):
    ensure_generated_dir()
    count = max(len(rows), 1)
    width = 1200
    row_h, gap, name_size = v16_fit_row_metrics(count, "standings")
    header_h = 58
    content_h = header_h + 26 + count * row_h + max(0, count - 1) * gap
    height = max(880, 245 + content_h + 180)
    img, draw = design_canvas("group_standing_template.png" if use_template else None, width, height, "purple")
    draw_design_header(draw, width, "مونديال المصيف 2026", clean_group_title_for_design(group_title), img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=112, accent="#22C55E")

    y = fy1 + 36
    rounded_rect(draw, (92, y, width-92, y+header_h), radius=22, fill="#05070D", outline="#FFFFFF66", width=1)
    draw_text(draw, (960, y+header_h//2), "المنتخب", get_font(26), fill="#FFFFFF")
    draw_text(draw, (500, y+header_h//2), "لعب", get_font(24), fill="#FDE68A")
    draw_text(draw, (370, y+header_h//2), "+/-", get_font(24), fill="#FDE68A")
    draw_text(draw, (225, y+header_h//2), "نقاط", get_font(24), fill="#FDE68A")
    y += header_h + 26

    for i, (team, played, diff, pts) in enumerate(rows, start=1):
        accent = v16_accent(i-1)
        rounded_rect(draw, (92, y, width-92, y+row_h), radius=26, fill="#0B1020", outline=accent, width=2)
        cy = y + row_h//2

        # V17: الرقم داخل أكثر من الحافة
        draw_text(draw, (1062, cy), str(i), get_font(max(28, name_size)), fill="#FDE68A" if i == 1 else "#FFFFFF")

        # V17: تكبير أعلام الترتيب
        flag_w = min(120, max(76, row_h-18))
        paste_flag(img, team, (900, cy-flag_w//2, 900+flag_w, cy+flag_w//2))

        draw_text(draw, (710, cy), team, get_font(name_size), fill="#FFFFFF", max_width=330)
        draw_text(draw, (500, cy), str(played), get_font(max(26, name_size-2)), fill="#FFFFFF")
        draw_text(draw, (370, cy), f"{diff:+d}", get_font(max(26, name_size-2)), fill="#FFFFFF")
        draw_text(draw, (225, cy), str(pts), get_font(max(30, name_size+3)), fill="#FDE68A")
        y += row_h + gap

    footer_event(draw, width, height)
    suffix = "template" if use_template else "auto"
    path = os.path.join(GENERATED_DIR, f"group_{suffix}_{_safe_filename(group_title)}.png")
    img.save(path, quality=95)
    return path

def create_top_scorers_template_image(items, use_template=True):
    ensure_generated_dir()
    items = sorted(items, key=lambda x: (-x[1], x[0]))
    count = max(len(items[:10]), 1)
    width = 1200
    row_h, gap, name_size = v16_fit_row_metrics(count, "scorers")
    content_h = count * row_h + max(0, count - 1) * gap
    height = max(830, 245 + content_h + 180)
    img, draw = design_canvas("scorers_template.png" if use_template else None, width, height, "gold")
    draw_design_header(draw, width, "مونديال المصيف 2026", "هدافين البطولة حتى الآن", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=112, accent="#F59E0B")

    available_h = (fy2 - fy1) - 70
    y = fy1 + 38 + max(0, (available_h - content_h) // 2)
    for i, (name, goals, team) in enumerate(items[:10], start=1):
        accent = "#F59E0B" if i == 1 else v16_accent(i)
        fill = "#1A1407" if i == 1 else "#0B1020"
        rounded_rect(draw, (92, y, width-92, y+row_h), radius=26, fill=fill, outline=accent, width=2)
        cy = y + row_h//2

        # V17: الرقم داخل أكثر من الحافة
        draw_text(draw, (1062, cy), str(i), get_font(max(30, name_size+4)), fill="#FDE68A")

        # V17: تكبير العلم
        if team:
            flag_w = min(110, max(68, row_h-18))
            paste_flag(img, team, (885, cy-flag_w//2, 885+flag_w, cy+flag_w//2))

        # مساحة أكبر للاسم
        draw_text(draw, (655, cy), name, get_font(name_size), fill="#FFFFFF", max_width=460)
        draw_text(draw, (245, cy), f"{goals} {'هدف' if goals == 1 else 'أهداف'}", get_font(max(24, name_size-2)), fill="#FDE68A")
        y += row_h + gap

    footer_event(draw, width, height)
    suffix = "template" if use_template else "auto"
    path = os.path.join(GENERATED_DIR, f"scorers_{suffix}.png")
    img.save(path, quality=95)
    return path

def build_group_standing_caption(group_title, rows):
    lines = [f"🏆 {group_title}", ""]
    for i, (team, played, diff, pts) in enumerate(rows, start=1):
        lines.append(f"{i}. {team} — لعب {played} | فارق {diff:+d} | نقاط {pts}")
    lines.append("\nالمصيف ينقل لكم الحدث")
    return "\n".join(lines)

async def design_matches_template_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day_name, matches = parse_matches_text(update.message.text)
    if not matches:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_مباريات\nالسادس\nفرنسا|البرازيل|10:00 م\nالأرجنتين|ألمانيا|12:00 ص")
        return
    try:
        path = create_matches_template_image(day_name, matches, use_template=True)
        await send_photo_path(update, path, build_design_matches_caption(day_name, matches))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم المباريات بالقالب ❌\n{e}")

async def design_matches_auto_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day_name, matches = parse_matches_text(update.message.text)
    if not matches:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_مباريات_تلقائي\nالسادس\nفرنسا|البرازيل|10:00 م")
        return
    try:
        path = create_matches_template_image(day_name, matches, use_template=False)
        await send_photo_path(update, path, build_design_matches_caption(day_name, matches))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم المباريات التلقائي ❌\n{e}")

async def design_match_results_template_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day_name, results = parse_match_results_design_text(update.message.text)
    if not results:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_نتائج_مباريات\nالسادس\nالسعودية|2|1|إسبانيا\nفرنسا|3|0|البرازيل")
        return
    try:
        path = create_match_results_template_image(day_name, results, use_template=True)
        await send_photo_path(update, path, build_match_results_caption(results))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم نتائج المباريات بالقالب ❌\n{e}")

async def design_match_results_auto_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day_name, results = parse_match_results_design_text(update.message.text)
    if not results:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_نتائج_مباريات_تلقائي\nالسادس\nالسعودية|2|1|إسبانيا")
        return
    try:
        path = create_match_results_template_image(day_name, results, use_template=False)
        await send_photo_path(update, path, build_match_results_caption(results))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم نتائج المباريات التلقائي ❌\n{e}")

async def design_group_standing_template_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    group_title, rows = parse_group_standing_text(update.message.text)
    if not rows:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_ترتيب_مجموعة\nالمجموعة C\nاسكتلندا|1|1|3\nالبرازيل|1|0|1\nالمغرب|1|0|1\nهايتي|1|-1|0")
        return
    try:
        path = create_group_standing_image(group_title, rows, use_template=True)
        await send_photo_path(update, path, build_group_standing_caption(group_title, rows))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم ترتيب المجموعة بالقالب ❌\n{e}")

async def design_group_standing_auto_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    group_title, rows = parse_group_standing_text(update.message.text)
    if not rows:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_ترتيب_مجموعة_تلقائي\nالمجموعة C\nاسكتلندا|1|1|3\nالبرازيل|1|0|1")
        return
    try:
        path = create_group_standing_image(group_title, rows, use_template=False)
        await send_photo_path(update, path, build_group_standing_caption(group_title, rows))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم ترتيب المجموعة التلقائي ❌\n{e}")

async def design_scorers_template_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    items = parse_scorers_text(update.message.text)
    if not items:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_هدافين\nراؤول خيمينيز|4|المكسيك\nفلوريان فيرتز|3|ألمانيا")
        return
    try:
        path = create_top_scorers_template_image(items, use_template=True)
        await send_photo_path(update, path, build_top_scorers_caption(items))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم الهدافين بالقالب ❌\n{e}")

async def design_scorers_auto_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    items = parse_scorers_text(update.message.text)
    if not items:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_هدافين_تلقائي\nراؤول خيمينيز|4|المكسيك")
        return
    try:
        path = create_top_scorers_template_image(items, use_template=False)
        await send_photo_path(update, path, build_top_scorers_caption(items))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم الهدافين التلقائي ❌\n{e}")


# ============================================================
# V23 إضافات التصميم الإضافي:
# ستايل 2 أزرق/ذهبي + ستايل 3 إطار + جميع المجموعات
# ============================================================

def _style2_canvas(width, height):
    img = Image.new("RGB", (width, height), "#061633")
    d = ImageDraw.Draw(img)
    for y in range(height):
        t = y / max(height, 1)
        d.line((0, y, width, y), fill=(3, int(24 + 20*t), int(70 + 30*t)))

    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.ellipse((80, -210, width-80, 560), fill=(15, 92, 190, 92))
    od.ellipse((width*0.55, 330, width*1.20, 1050), fill=(0, 120, 255, 46))
    od.rectangle((width*0.61, 430, width*0.90, 640), fill=(0, 115, 255, 36))
    od.ellipse((-270, 500, 300, height+120), fill=(0, 70, 180, 70))
    overlay = overlay.filter(ImageFilter.GaussianBlur(28))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    d = ImageDraw.Draw(img)

    for i in range(0, 64, 7):
        d.arc((80+i*2, 20+i, width-80-i*2, 720+i), 190, 335, fill="#1D6FEA55", width=2)
    for i in range(18):
        x = 60 + i * 18
        y = height - 215 + (i % 3) * 8
        d.ellipse((x, y, x+5, y+5), fill="#FBBF2435")
    return img, d

def _ampm_from_time(t):
    t = normalize_name(t)
    if "ص" in t or "AM" in t.upper():
        period = "ص"
    elif "م" in t or "PM" in t.upper():
        period = "م"
    else:
        period = ""
    tm = re.sub(r"(ص|م|AM|PM|am|pm)", "", t).strip()
    return tm, period

def create_matches_style2_image(day_name, matches):
    ensure_generated_dir()
    count = max(len(matches), 1)
    width = 1200

    if count == 1:
        row_h, gap, header_h, footer_h, card_pad = 245, 0, 330, 95, 95
    elif count == 2:
        row_h, gap, header_h, footer_h, card_pad = 215, 28, 320, 95, 115
    elif count <= 4:
        row_h, gap, header_h, footer_h, card_pad = 175, 22, 310, 95, 160
    else:
        row_h, gap, header_h, footer_h, card_pad = 150, 15, 290, 90, 150

    height = max(820, header_h + count * row_h + max(0, count-1) * gap + footer_h)
    img, draw = _style2_canvas(width, height)

    draw_text(draw, (width//2, 62), "MONDIAL AL MASEEF 2026", get_font(30), fill="#FFFFFF")
    draw_text(draw, (width//2, 132), "مباريات اليوم", get_font(80 if count <= 2 else 72), fill="#FFFFFF")
    rounded_rect(draw, (width//2-235, 202, width//2+235, 260), radius=18, fill="#FBBF24", outline="#00000055", width=2)
    draw_text(draw, (width//2, 232), f"اليوم {day_name}", get_font(30), fill="#061633")

    y = header_h
    if count == 1:
        y = 350

    for a, b, t in matches:
        x1, x2 = card_pad, width-card_pad
        rounded_rect(draw, (x1, y, x2, y+row_h), radius=32, fill="#0638A5", outline="#14B8F5", width=4)

        # نفس روح الأطراف الملونة
        draw.line((x1+14, y+row_h-4, x2-14, y+row_h-4), fill="#EF4444", width=4)
        draw.line((x1+170, y+6, x2-170, y+6), fill="#22C55E", width=3)
        draw.line((x1+390, y+6, x1+470, y+6), fill="#FBBF24", width=3)
        draw.arc((x1, y, x1+58, y+58), 180, 270, fill="#14B8F5", width=8)
        draw.arc((x2-58, y+row_h-58, x2, y+row_h), 0, 90, fill="#EF4444", width=8)

        cy = y + row_h//2
        flag_w = min(158 if count <= 2 else 132, row_h-48)

        paste_flag(img, a, (x1+42, cy-flag_w//2, x1+42+flag_w, cy+flag_w//2))
        paste_flag(img, b, (x2-42-flag_w, cy-flag_w//2, x2-42, cy+flag_w//2))

        name_font = get_font(34 if count <= 2 else 27)
        draw_text(draw, (x1+42+flag_w//2, y+row_h-38), a, name_font, fill="#FFFFFF", max_width=260)
        draw_text(draw, (x2-42-flag_w//2, y+row_h-38), b, name_font, fill="#FFFFFF", max_width=260)

        tm, period = _ampm_from_time(t)
        time_font = get_font(68 if count == 1 else (60 if count == 2 else 52))
        draw_text(draw, (width//2, cy-18), tm, time_font, fill="#FBBF24")
        if period:
            draw_text(draw, (width//2, cy+46), period, get_font(36), fill="#FBBF24")
        elif count == 1:
            draw_text(draw, (width//2, cy+48), "بتوقيت السعودية", get_font(26), fill="#E5E7EB")

        y += row_h + gap

    draw_text(draw, (width//2, height-44), "المصيف ينقل لكم الحدث", get_font(30), fill="#FBBF24")
    path = os.path.join(GENERATED_DIR, f"matches_style2_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path

def create_match_results_style2_image(day_name, results):
    ensure_generated_dir()
    count = max(len(results), 1)
    width = 1200

    if count == 1:
        row_h, gap, header_h, footer_h, card_pad = 245, 0, 330, 95, 95
    elif count == 2:
        row_h, gap, header_h, footer_h, card_pad = 215, 28, 320, 95, 115
    elif count <= 4:
        row_h, gap, header_h, footer_h, card_pad = 175, 22, 310, 95, 160
    else:
        row_h, gap, header_h, footer_h, card_pad = 150, 15, 290, 90, 150

    height = max(820, header_h + count * row_h + max(0, count-1) * gap + footer_h)
    img, draw = _style2_canvas(width, height)

    draw_text(draw, (width//2, 62), "MONDIAL AL MASEEF 2026", get_font(30), fill="#FFFFFF")
    draw_text(draw, (width//2, 132), "نتائج اليوم", get_font(80 if count <= 2 else 72), fill="#FFFFFF")
    rounded_rect(draw, (width//2-235, 202, width//2+235, 260), radius=18, fill="#FBBF24", outline="#00000055", width=2)
    draw_text(draw, (width//2, 232), f"اليوم {day_name}", get_font(30), fill="#061633")

    y = 350 if count == 1 else header_h
    for a, sa, sb, b in results:
        x1, x2 = card_pad, width-card_pad
        rounded_rect(draw, (x1, y, x2, y+row_h), radius=32, fill="#0638A5", outline="#14B8F5", width=4)
        draw.line((x1+14, y+row_h-4, x2-14, y+row_h-4), fill="#EF4444", width=4)
        draw.line((x1+170, y+6, x2-170, y+6), fill="#22C55E", width=3)
        draw.arc((x1, y, x1+58, y+58), 180, 270, fill="#14B8F5", width=8)
        draw.arc((x2-58, y+row_h-58, x2, y+row_h), 0, 90, fill="#EF4444", width=8)

        cy = y + row_h//2
        flag_w = min(158 if count <= 2 else 132, row_h-48)
        paste_flag(img, a, (x1+42, cy-flag_w//2, x1+42+flag_w, cy+flag_w//2))
        paste_flag(img, b, (x2-42-flag_w, cy-flag_w//2, x2-42, cy+flag_w//2))

        name_font = get_font(34 if count <= 2 else 27)
        draw_text(draw, (x1+42+flag_w//2, y+row_h-38), a, name_font, fill="#FFFFFF", max_width=260)
        draw_text(draw, (x2-42-flag_w//2, y+row_h-38), b, name_font, fill="#FFFFFF", max_width=260)

        rounded_rect(draw, (width//2-125, cy-50, width//2+125, cy+50), radius=20, fill="#FBBF24", outline="#00000088", width=2)
        draw_text(draw, (width//2, cy), f"{sb} - {sa}", get_font(60 if count <= 2 else 52), fill="#061633")
        y += row_h + gap

    draw_text(draw, (width//2, height-44), "المصيف ينقل لكم الحدث", get_font(30), fill="#FBBF24")
    path = os.path.join(GENERATED_DIR, f"results_style2_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path

def create_scorers_style2_image(items):
    ensure_generated_dir()
    items = sorted(items, key=lambda x: (-x[1], x[0]))[:12]
    count = max(len(items), 1)
    width = 1200
    row_h = 105 if count <= 8 else 88
    gap = 14 if count <= 8 else 10
    height = max(900, 310 + count*row_h + (count-1)*gap + 90)
    img, draw = _style2_canvas(width, height)
    draw_text(draw, (width//2, 80), "MONDIAL AL MASEEF 2026", get_font(30), fill="#FFFFFF")
    draw_text(draw, (width//2, 150), "هدافين البطولة", get_font(72), fill="#FFFFFF")
    rounded_rect(draw, (width//2-210, 215, width//2+210, 268), radius=18, fill="#FBBF24", outline="#00000055", width=2)
    draw_text(draw, (width//2, 242), "تحديث مستمر", get_font(28), fill="#061633")
    y = 315
    for i, (name, goals, team) in enumerate(items, start=1):
        rounded_rect(draw, (120, y, width-120, y+row_h), radius=24, fill="#0638A5", outline="#14B8F5", width=3)
        cy = y + row_h//2
        draw_text(draw, (1030, cy), str(i), get_font(34), fill="#FBBF24")
        if team:
            paste_flag(img, team, (850, cy-38, 930, cy+38))
        draw_text(draw, (620, cy), name, get_font(34), fill="#FFFFFF", max_width=430)
        draw_text(draw, (240, cy), f"{goals} {'هدف' if goals == 1 else 'أهداف'}", get_font(30), fill="#FBBF24")
        y += row_h + gap
    draw_text(draw, (width//2, height-42), "المصيف ينقل لكم الحدث", get_font(28), fill="#FBBF24")
    path = os.path.join(GENERATED_DIR, "scorers_style2.png")
    img.save(path, quality=95)
    return path

def create_group_style2_image(group_title, rows):
    ensure_generated_dir()
    count = max(len(rows), 1)
    width = 1200
    row_h = 105 if count <= 6 else 86
    gap = 14 if count <= 6 else 10
    height = max(870, 310 + count*row_h + (count-1)*gap + 90)
    img, draw = _style2_canvas(width, height)
    draw_text(draw, (width//2, 80), "MONDIAL AL MASEEF 2026", get_font(30), fill="#FFFFFF")
    draw_text(draw, (width//2, 150), "ترتيب المجموعة", get_font(72), fill="#FFFFFF")
    rounded_rect(draw, (width//2-220, 215, width//2+220, 268), radius=18, fill="#FBBF24", outline="#00000055", width=2)
    draw_text(draw, (width//2, 242), clean_group_title_for_design(group_title), get_font(30), fill="#061633")
    y = 315
    for i, (team, played, diff, pts) in enumerate(rows, start=1):
        rounded_rect(draw, (120, y, width-120, y+row_h), radius=24, fill="#0638A5", outline="#14B8F5", width=3)
        cy = y + row_h//2
        draw_text(draw, (1040, cy), str(i), get_font(34), fill="#FBBF24")
        paste_flag(img, team, (870, cy-40, 950, cy+40))
        draw_text(draw, (690, cy), team, get_font(34), fill="#FFFFFF", max_width=330)
        draw_text(draw, (420, cy), f"لعب {played}", get_font(24), fill="#E5E7EB")
        draw_text(draw, (300, cy), f"{int(diff):+d}", get_font(28), fill="#E5E7EB")
        draw_text(draw, (190, cy), str(pts), get_font(34), fill="#FBBF24")
        y += row_h + gap
    draw_text(draw, (width//2, height-42), "المصيف ينقل لكم الحدث", get_font(28), fill="#FBBF24")
    path = os.path.join(GENERATED_DIR, f"group_style2_{_safe_filename(group_title)}.png")
    img.save(path, quality=95)
    return path

def parse_all_groups_text(text):
    groups = []
    current_title = None
    current_rows = []
    for raw in (text or "").splitlines()[1:]:
        line = raw.strip()
        if not line:
            continue
        if "|" not in line:
            if current_title and current_rows:
                groups.append((current_title, current_rows))
                current_rows = []
            current_title = line
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 4 and current_title:
            try:
                current_rows.append((parts[0], int(parts[1]), int(parts[2]), int(parts[3])))
            except Exception:
                pass
    if current_title and current_rows:
        groups.append((current_title, current_rows))
    return groups

def create_all_groups_image(groups):
    ensure_generated_dir()
    width, height = 1800, 2400
    img, draw = _style2_canvas(width, height)
    draw_text(draw, (width//2, 90), "MONDIAL AL MASEEF 2026", get_font(40), fill="#FFFFFF")
    draw_text(draw, (width//2, 165), "ترتيب جميع المجموعات", get_font(72), fill="#FFFFFF")
    cols = 3
    margin_x, gap_x = 80, 38
    card_w = (width - 2*margin_x - (cols-1)*gap_x) // cols
    card_h = 475
    start_y = 260
    gap_y = 42
    for idx, (title, rows) in enumerate(groups[:12]):
        c = idx % cols
        r = idx // cols
        x = margin_x + c*(card_w+gap_x)
        y = start_y + r*(card_h+gap_y)
        rounded_rect(draw, (x, y, x+card_w, y+card_h), radius=28, fill="#0638A5EE", outline="#14B8F5", width=3)
        rounded_rect(draw, (x+18, y+18, x+card_w-18, y+68), radius=18, fill="#FBBF24", outline="#00000055", width=1)
        draw_text(draw, (x+card_w//2, y+43), clean_group_title_for_design(title), get_font(26), fill="#061633")
        yy = y + 92
        for pos, (team, played, diff, pts) in enumerate(rows[:4], start=1):
            rounded_rect(draw, (x+20, yy, x+card_w-20, yy+72), radius=16, fill="#061633AA", outline="#FFFFFF22", width=1)
            cy = yy + 36
            draw_text(draw, (x+card_w-45, cy), str(pos), get_font(22), fill="#FBBF24")
            paste_flag(img, team, (x+card_w-125, cy-24, x+card_w-75, cy+24))
            draw_text(draw, (x+card_w-240, cy), team, get_font(22), fill="#FFFFFF", max_width=190)
            draw_text(draw, (x+155, cy), str(pts), get_font(26), fill="#FBBF24")
            draw_text(draw, (x+75, cy), f"{diff:+d}", get_font(21), fill="#E5E7EB")
            yy += 84
    draw_text(draw, (width//2, height-70), "المصيف ينقل لكم الحدث", get_font(36), fill="#FBBF24")
    path = os.path.join(GENERATED_DIR, "all_groups_style2.png")
    img.save(path, quality=95)
    return path

def _frame_canvas(width, height):
    img = Image.new("RGB", (width, height), "#07111B")
    d = ImageDraw.Draw(img)
    for y in range(height):
        t = y / max(height, 1)
        d.line((0, y, width, y), fill=(5, int(18+20*t), int(30+24*t)))
    overlay = Image.new("RGBA", (width, height), (0,0,0,0))
    od = ImageDraw.Draw(overlay)
    od.ellipse((-200, -120, 440, 460), fill=(10,200,120,30))
    od.ellipse((width-360, height-420, width+260, height+160), fill=(70,90,255,40))
    img = Image.alpha_composite(img.convert("RGBA"), overlay.filter(ImageFilter.GaussianBlur(40))).convert("RGB")
    return img, ImageDraw.Draw(img)

def create_match_frame_style_image(day_name, items, is_results=False):
    ensure_generated_dir()
    count = max(len(items), 1)
    width = 1200
    row_h = 126 if count <= 6 else 104
    gap = 18 if count <= 6 else 12
    height = max(620, 220 + count*row_h + (count-1)*gap + 105)
    img, draw = _frame_canvas(width, height)
    draw_text(draw, (width//2, 80), "مونديال المصيف 2026", get_font(50), fill="#FFFFFF")
    draw_text(draw, (width//2, 140), f"{'نتائج' if is_results else 'مباريات'} اليوم {day_name}", get_font(34), fill="#FDE68A")
    y = 215
    for item in items:
        if is_results:
            a, sa, sb, b = item
            center = f"{sb} - {sa}"
        else:
            a, b, t = item
            center = t
        x1, x2 = 120, width-120
        rounded_rect(draw, (x1, y, x2, y+row_h), radius=28, fill="#07110FCC", outline="#22C55E", width=3)
        draw.line((x1+5, y+row_h-5, x1+260, y+row_h-5), fill="#F97316", width=10)
        draw.line((x2-260, y+row_h-5, x2-5, y+row_h-5), fill="#2563EB", width=10)
        draw.line((x1+10, y+5, x2-10, y+5), fill="#22C55E", width=4)
        cy = y + row_h//2
        flag_w = min(84, row_h-36)
        paste_flag(img, a, (x2-160, cy-flag_w//2, x2-160+flag_w, cy+flag_w//2))
        paste_flag(img, b, (x1+80, cy-flag_w//2, x1+80+flag_w, cy+flag_w//2))
        draw_text(draw, (x2-320, cy), a, get_font(30), fill="#FFFFFF", max_width=270)
        draw_text(draw, (x1+280, cy), b, get_font(30), fill="#FFFFFF", max_width=270)
        rounded_rect(draw, (width//2-95, cy-38, width//2+95, cy+38), radius=18, fill="#B8FFF0", outline="#FFFFFFAA", width=2)
        draw_text(draw, (width//2, cy), center, get_font(34), fill="#061633")
        y += row_h + gap
    draw_text(draw, (width//2, height-42), "المصيف ينقل لكم الحدث", get_font(28), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, f"{'results' if is_results else 'matches'}_frame_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path



def build_design_matches_caption(day_name, matches):
    lines = [
        "🏆 مونديال المصيف 2026 🏆",
        f"🔥 مباريات اليوم ( {day_name} ) 🔥",
        ""
    ]
    for a, b, t in matches:
        lines.append(f"{a} × {b} — {t}")
    lines.append("")
    lines.append("المصيف ينقل لكم الحدث")
    return "\n".join(lines)


async def design_matches_style2_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        day_name, matches = parse_matches_text(update.message.text)
        if not matches:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_مباريات_ستايل2\nالسابع\nالبرتغال|الكونغو الديمقراطية|8:00 م")
            return
        path = create_matches_style2_image(day_name, matches)
        await send_photo_path(update, path, build_design_matches_caption(day_name, matches))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم مباريات ستايل2 ❌\n{e}")

async def design_match_results_style2_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        day_name, results = parse_match_results_design_text(update.message.text)
        if not results:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_نتائج_مباريات_ستايل2\nالسابع\nالبرتغال|2|1|الكونغو الديمقراطية")
            return
        path = create_match_results_style2_image(day_name, results)
        await send_photo_path(update, path, build_match_results_caption(results))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم نتائج ستايل2 ❌\n{e}")

async def design_scorers_style2_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        items = parse_scorers_text(update.message.text)
        if not items:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_هدافين_ستايل2\nراؤول خيمينيز|4|المكسيك")
            return
        path = create_scorers_style2_image(items)
        await send_photo_path(update, path, build_top_scorers_caption(items))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم الهدافين ستايل2 ❌\n{e}")

async def design_group_style2_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        group_title, rows = parse_group_standing_text(update.message.text)
        if not rows:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_ترتيب_مجموعة_ستايل2\nالمجموعة C\nالبرازيل|1|0|1")
            return
        path = create_group_style2_image(group_title, rows)
        await send_photo_path(update, path, build_group_standing_caption(group_title, rows))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم ترتيب المجموعة ستايل2 ❌\n{e}")

async def design_matches_frame_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        day_name, matches = parse_matches_text(update.message.text)
        if not matches:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_مباريات_اطار\nالسابع\nالبرتغال|الكونغو الديمقراطية|8:00 م")
            return
        path = create_match_frame_style_image(day_name, matches, False)
        await send_photo_path(update, path, build_design_matches_caption(day_name, matches))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم مباريات الإطار ❌\n{e}")

async def design_results_frame_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        day_name, results = parse_match_results_design_text(update.message.text)
        if not results:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_نتائج_مباريات_اطار\nالسابع\nالبرتغال|2|1|الكونغو الديمقراطية")
            return
        path = create_match_frame_style_image(day_name, results, True)
        await send_photo_path(update, path, build_match_results_caption(results))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم نتائج الإطار ❌\n{e}")

async def design_all_groups_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        groups = parse_all_groups_text(update.message.text)
        if not groups:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_جميع_المجموعات\nالمجموعة A\nفريق|1|0|3\n...\nالمجموعة B\nفريق|1|0|3")
            return
        path = create_all_groups_image(groups)
        await send_photo_path(update, path, "ترتيب جميع المجموعات ✅")
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم جميع المجموعات ❌\n{e}")


def main():
    if not TOKEN:
        raise RuntimeError("ضع توكن البوت في متغير البيئة BOT_TOKEN")
    ensure_flags_assets()
    ensure_design_assets()
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/start"), start))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/(?:من_انا|معرفي)"), who_am_i))
    app.add_handler(MessageHandler(filters.Document.ALL, remember_last_file))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تفعيل_الصور_التلقائية"), admin_only(enable_auto_images)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إيقاف_الصور_التلقائية"), admin_only(disable_auto_images)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/ايقاف_الصور_التلقائية"), admin_only(disable_auto_images)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صورة_اليوم"), daily_image_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صورة_الترتيب"), overall_image_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صورة_الاساطير"), legends_image_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صورة_احصائيات"), dashboard_sheet_image_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صور_الاحصائيات"), all_dashboard_images_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/بطاقة"), participant_card_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تقرير_الفترة"), period_report_command))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعلان_اليوم"), announcement_day_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/ملخص_اليوم"), summary_day_command))
    # V23 أوامر التصميم الإضافي
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_مباريات_ستايل2(?:\s|$)"), admin_only(design_matches_style2_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_نتائج_مباريات_ستايل2(?:\s|$)"), admin_only(design_match_results_style2_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_هدافين_ستايل2(?:\s|$)"), admin_only(design_scorers_style2_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_ترتيب_مجموعة_ستايل2(?:\s|$)"), admin_only(design_group_style2_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_مباريات_اطار(?:\s|$)"), admin_only(design_matches_frame_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_نتائج_مباريات_اطار(?:\s|$)"), admin_only(design_results_frame_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_جميع_المجموعات(?:\s|$)"), admin_only(design_all_groups_command)))

    # أوامر التصاميم بالقالب والتلقائي — الأوامر الأطول قبل الأقصر
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_مباريات_تلقائي(?:\s|$)"), admin_only(design_matches_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_نتائج_مباريات_تلقائي(?:\s|$)"), admin_only(design_match_results_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_ترتيب_مجموعة_تلقائي(?:\s|$)"), admin_only(design_group_standing_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_هدافين_تلقائي(?:\s|$)"), admin_only(design_scorers_auto_command)))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_مباريات(?:\s|$)"), admin_only(design_matches_template_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_نتائج_مباريات(?:\s|$)"), admin_only(design_match_results_template_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_ترتيب_مجموعة(?:\s|$)"), admin_only(design_group_standing_template_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_هدافين(?:\s|$)"), admin_only(design_scorers_template_command)))

    # اسم قديم للتوافق فقط
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج_مباريات_اليوم(?:\s|$)"), admin_only(design_match_results_template_command)))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استيراد_ملف"), admin_only(import_excel_file)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعتماد_استيراد"), admin_only(approve_import)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إلغاء_استيراد"), admin_only(cancel_import)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الغاء_استيراد"), admin_only(cancel_import)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نسخة_احتياطية"), admin_only(backup_zip)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استرجاع_نسخة"), admin_only(restore_backup_zip)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تنظيف_الأيام"), admin_only(clean_days)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تنظيف_الايام"), admin_only(clean_days)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تنظيف_الملفات"), admin_only(clean_temp_files)))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/بدء_الكاس(?:\s|$)"), admin_only(start_cup_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/حالة_الكاس(?:\s|$)"), admin_only(cup_status_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج_الكاس(?:\s|$)"), admin_only(cup_results_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مواجهات_الكاس(?:\s|$)"), admin_only(cup_matches_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إعادة_الكاس_من(?:\s|$)"), admin_only(reset_cup_from_day_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعادة_الكاس_من(?:\s|$)"), admin_only(reset_cup_from_day_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الغاء_الكاس(?:\s|$)"), admin_only(cancel_cup_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إلغاء_الكاس(?:\s|$)"), admin_only(cancel_cup_command)))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اضافه"), admin_only(add_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعتماد_نتائج(?:\s|$)"), admin_only(approve_results_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج"), admin_only(results_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الترتيب_العام"), overall))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/ترتيب_نص"), ranking_text))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/?(?:احصائيات|إحصائيات)(?:\s|$)"), dashboard))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/(الأيام|الايام)"), list_days))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/فحص"), inspect_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مشاركين"), participants_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اسطورة"), legend_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مقارنة"), compare_days))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_الكل"), admin_only(clear_all)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_يوم"), admin_only(clear_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_نتائج"), admin_only(clear_results)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استرجاع_آخر"), admin_only(restore_last)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/قفل_يوم"), admin_only(lock_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/فتح_يوم"), admin_only(unlock_day)))
    app.run_polling()



# ===== V22 patched overrides =====

def _clean_display_name(text):
    text = normalize_name(text)
    # إزالة الرموز الشائعة والإيموجي حتى لا تظهر مربعات في الصور
    text = re.sub(r"[\U00010000-\U0010ffff]", "", text)
    text = re.sub(r"[⚽🏆👑🧤🔥✅❌⭐🎯🥇🥈🥉📌📊⚔️😅]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "-"


def build_daily_summary_text(day):
    data = build_day_summary(day)
    winners = " + ".join(data["winners"]) if data["winners"] else "لا يوجد"
    top_rows = data["sorted_rows"][:5]
    lines = [
        f"📊 ملخص اليوم {ordinal_day(day)}",
        f"👥 عدد المشاركين: {len(data['participants'])}",
        f"🏆 أسطورة اليوم: {winners} — {data['max_score']} نقطة",
        "",
        "🔥 أفضل 5:",
    ]
    for idx, r in enumerate(top_rows, start=1):
        lines.append(f"{idx}. {r['participant']} — {r['total']} نقطة")
    lines.append("")
    lines.extend(matchup_lines(day))
    return "\n".join(lines)


def create_daily_result_image(day, goals_count=None, clean_sheets=None):
    ensure_generated_dir()
    data = build_day_summary(day)
    rows = data["sorted_rows"]
    matchups = generate_matchups_for_day(day)
    match_rows = len(matchups.get("pairs", [])) + (1 if matchups.get("bye") else 0)
    goals_lines = [f"{_clean_display_name(p)} — {c}" for p, c in (goals_count or {}).items()] if goals_count else [f"{_clean_display_name(p)} — {pts} نقطة" for p, pts in data["player_points"].most_common(5)]
    clean_lines = [ _clean_display_name(x) for x in (clean_sheets or list(data["keeper_points"].keys()) or ["لا يوجد"]) ]
    goals_lines = goals_lines or ["لا يوجد"]
    clean_lines = clean_lines or ["لا يوجد"]

    width = 1400
    base_h = 420 + len(rows) * 56 + max(170, 85 + match_rows * 40) + 250
    height = max(1220, base_h)
    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, f"فانتزي المصيف 2026 - اليوم {ordinal_day(day)}", "أسطورة اليوم وترتيب المشاركين", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=110, accent="#8B5CF6")

    legends = " + ".join(data["winners"]) if data["winners"] else "لا يوجد"
    rounded_rect(draw, (90, fy1 + 18, 1010, fy1 + 150), radius=32, fill="#7C3AEDDD", outline="#FFFFFF40", width=2)
    draw_text(draw, (550, fy1 + 58), "أسطورة اليوم", get_font(28), fill="#FFFFFF")
    draw_text(draw, (550, fy1 + 112), f"{legends} - {data['max_score']} نقطة", get_font(42), fill="#FFF6D6", max_width=850)
    rounded_rect(draw, (1040, fy1 + 18, 1310, fy1 + 150), radius=26, fill="#F59E0BDD", outline="#FFFFFF33", width=2)
    draw_text(draw, (1175, fy1 + 58), "المشاركون", get_font(24), fill="#FFFFFF")
    draw_text(draw, (1175, fy1 + 112), f"{len(data['participants'])}", get_font(46), fill="#FFFFFF")

    y = fy1 + 185
    rounded_rect(draw, (90, y, 1310, y + 56), radius=18, fill="#05070D", outline="#FFFFFF55", width=1)
    headers = [(1180, "المشارك"), (870, "النقاط"), (655, "الحارس"), (430, "الكابتن"), (190, "المركز")]
    for x, t in headers:
        draw_text(draw, (x, y + 28), t, get_font(22), fill="#FDE68A")
    y += 70
    row_h = 50
    for idx, r in enumerate(rows, start=1):
        accent = "#F59E0B" if idx == 1 else "#FFFFFF22"
        fill = "#1A1407" if idx == 1 else ("#0B1020" if idx % 2 else "#10172A")
        rounded_rect(draw, (90, y, 1310, y + row_h), radius=16, fill=fill, outline=accent, width=2 if idx == 1 else 1)
        cy = y + row_h // 2
        draw_text(draw, (1180, cy), _clean_display_name(r["participant"]), get_font(24), fill="#FFFFFF", max_width=300)
        draw_text(draw, (870, cy), str(r["total"]), get_font(28), fill="#FDE68A")
        draw_text(draw, (655, cy), f"+{r['keeper_points']}", get_font(25), fill="#A7F3D0")
        draw_text(draw, (430, cy), f"+{r['captain_points']}", get_font(25), fill="#C4B5FD")
        medal = "🥇" if idx == 1 else "🥈" if idx == 2 else "🥉" if idx == 3 else str(idx)
        draw_text(draw, (190, cy), medal, get_font(24), fill="#FFFFFF")
        y += row_h + 8

    match_box_h = max(170, 85 + match_rows * 38)
    rounded_rect(draw, (90, y + 20, 1310, y + 20 + match_box_h), radius=26, fill="#091122DD", outline="#22C55E", width=2)
    draw_text(draw, (700, y + 56), "مواجهات اليوم", get_font(28), fill="#FDE68A")
    yy = y + 100
    if not matchups.get("pairs") and not matchups.get("bye"):
        draw_text(draw, (700, yy), "لا توجد مواجهات — لا يوجد مشاركون", get_font(23), fill="#FFFFFF")
        yy += 36
    else:
        for p in matchups.get("pairs", []):
            line = f"{_clean_display_name(p['a'])} {p['a_points']} - {p['b_points']} {_clean_display_name(p['b'])} - الفائز: {_clean_display_name(p['winner'])}"
            draw_text(draw, (700, yy), line, get_font(22), fill="#FFFFFF", max_width=1100)
            yy += 36
        if matchups.get("bye"):
            draw_text(draw, (700, yy), f"راحة الجولة: {_clean_display_name(matchups['bye'])}", get_font(22), fill="#FDE68A")
            yy += 36
    y = y + 20 + match_box_h + 28

    box_h = 185
    rounded_rect(draw, (90, y, 675, y + box_h), radius=26, fill="#091122DD", outline="#A855F7", width=2)
    rounded_rect(draw, (725, y, 1310, y + box_h), radius=26, fill="#091122DD", outline="#06B6D4", width=2)
    draw_text(draw, (382, y + 38), "الهدافين", get_font(30), fill="#FFFFFF")
    draw_text(draw, (382, y + 112), "\n".join(goals_lines[:5]), get_font(24), fill="#E5E7EB", max_width=500)
    draw_text(draw, (1018, y + 38), "الكلين شيت", get_font(30), fill="#FFFFFF")
    draw_text(draw, (1018, y + 112), "\n".join(clean_lines[:5]), get_font(24), fill="#E5E7EB", max_width=500)

    footer_event(draw, width, height)
    path = os.path.join(GENERATED_DIR, f"daily_result_day_{day}.png")
    img.save(path, quality=95)
    return path


def create_legends_image(start_day=1, end_day=31):
    ensure_generated_dir()
    stats = collect_stats(start_day, end_day)
    days = stats["days"]
    width = 1200
    count = max(len(days), 1)
    row_h = 72
    gap = 12
    content_h = 86 + count * row_h + max(0, count - 1) * gap
    height = max(820, 250 + content_h + 140)
    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, "سجل أساطير الفانتزي", f"من اليوم {start_day} إلى اليوم {end_day}", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=220, bottom_pad=110, accent="#8B5CF6")

    y = fy1 + 24
    rounded_rect(draw, (92, y, width-92, y+56), radius=18, fill="#05070D", outline="#FFFFFF55", width=1)
    draw_text(draw, (1010, y+28), "اليوم", get_font(22), fill="#FDE68A")
    draw_text(draw, (650, y+28), "أسطورة اليوم", get_font(22), fill="#FDE68A")
    draw_text(draw, (220, y+28), "النقاط", get_font(22), fill="#FDE68A")
    y += 72

    if not days:
        rounded_rect(draw, (92, y, width-92, y+70), radius=18, fill="#0B1020", outline="#FFFFFF22", width=1)
        draw_text(draw, (width//2, y+35), "لا توجد أيام في هذا النطاق", get_font(24), fill="#FFFFFF")
    else:
        for idx, day in enumerate(days, start=1):
            info = stats["per_day"].get(day, {})
            winners = " + ".join(_clean_display_name(w) for w in info.get("winners", [])) if info.get("winners") else "لا يوجد"
            max_score = info.get("max_score", 0)
            accent = "#F59E0B" if idx == 1 else v16_accent(idx)
            rounded_rect(draw, (92, y, width-92, y+row_h), radius=20, fill="#0B1020", outline=accent, width=2)
            cy = y + row_h//2
            draw_text(draw, (1010, cy), f"اليوم {ordinal_day(day)}", get_font(24), fill="#FFFFFF")
            draw_text(draw, (650, cy), winners, get_font(24), fill="#FDE68A", max_width=470)
            draw_text(draw, (220, cy), f"{max_score} نقطة", get_font(24), fill="#A7F3D0")
            y += row_h + gap

    footer_event(draw, width, height)
    path = os.path.join(GENERATED_DIR, f"legends_{start_day}_{end_day}.png")
    img.save(path, quality=95)
    return path


def _rank_map_until(day):
    if day is None or day < 1:
        return {}
    stats = collect_stats(1, day)
    return {name: idx + 1 for idx, name in enumerate(stats.get("ranking", []))}


def _biggest_daily_score(start_day, end_day):
    best_name, best_score, best_day = "-", -1, None
    for day in get_existing_days(start_day, end_day):
        for r in read_day_rows(day):
            if r.get("participated") and r.get("total", 0) > best_score:
                best_name, best_score, best_day = r["participant"], r.get("total", 0), day
    if best_score < 0:
        return "-", 0, "-"
    return best_name, best_score, best_day


def _movement_text(name, old_rank, new_rank):
    if not name or old_rank is None or new_rank is None or old_rank == 999 or new_rank == 999:
        return "-"
    return f"{_clean_display_name(name)}\nمن المركز {old_rank} إلى {new_rank}"


def create_participant_card_image(name, start_day=1, end_day=31):
    ensure_generated_dir()
    stats = collect_stats(start_day, end_day)
    if name not in PARTICIPANTS:
        raise ValueError("اسم المشارك غير موجود")
    rank = stats["ranking"].index(name) + 1 if name in stats["ranking"] else "-"
    total = stats["totals"].get(name, 0)
    day_scores = stats["scores_by_day"].get(name, {})
    best_day = max(day_scores, key=lambda d: day_scores[d]) if day_scores else "-"
    worst_day = min(day_scores, key=lambda d: day_scores[d]) if day_scores else "-"
    best_score = day_scores.get(best_day, 0) if best_day != "-" else 0
    worst_score = day_scores.get(worst_day, 0) if worst_day != "-" else 0
    best_cap, best_cap_pts, best_cap_day = participant_best_captain(name, start_day, end_day)
    wins = matchup_wins_map(start_day, end_day).get(name, 0)
    pc = stats["participation_count"].get(name, 0)
    days_count = len(stats["days"])
    pct = f"{round(pc / days_count * 100, 1)}%" if days_count else "0%"

    width, height = 1200, 920
    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, "بطاقة مشارك فانتزي المصيف", _clean_display_name(name), img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=225, bottom_pad=110, accent="#8B5CF6")
    cards = [
        (90, fy1+20, 360, fy1+145, "المركز", f"#{rank}"),
        (465, fy1+20, 735, fy1+145, "النقاط", str(total)),
        (840, fy1+20, 1110, fy1+145, "أسطورة اليوم", str(stats["daily_wins"].get(name,0))),
        (90, fy1+185, 360, fy1+310, "أفضل يوم", f"{best_day} — {best_score} نقاط"),
        (465, fy1+185, 735, fy1+310, "أسوأ يوم", f"{worst_day} — {worst_score} نقاط"),
        (840, fy1+185, 1110, fy1+310, "نسبة المشاركة", pct),
        (90, fy1+350, 545, fy1+490, "أفضل كابتن", f"{_clean_display_name(best_cap)}\n+{best_cap_pts} نقطة"),
        (655, fy1+350, 1110, fy1+490, "انتصارات المواجهات اليومية", str(wins)),
    ]
    for x1,y1,x2,y2,t,v in cards:
        rounded_rect(draw,(x1,y1,x2,y2), radius=28, fill="#091122DD", outline="#FFFFFF25", width=2)
        draw_text(draw, ((x1+x2)//2, y1+34), t, get_font(26), fill="#E5E7EB")
        draw_text(draw, ((x1+x2)//2, y1+88), v, get_font(34), fill="#FFFFFF", max_width=x2-x1-28)
    footer_event(draw, width, height)
    path = os.path.join(GENERATED_DIR, f"participant_card_{_safe_filename(name)}.png")
    img.save(path, quality=95)
    return path


def create_period_report_image(start_day, end_day):
    ensure_generated_dir()
    days, totals, cap_points, keeper_points, player_impact, player_zero, active_counts = period_stats(start_day, end_day)
    if not days:
        raise ValueError("ما فيه أيام في الفترة")

    stats_range = collect_stats(start_day, end_day)
    champion = totals.most_common(1)[0][0] if totals else "-"
    champ_points = totals.get(champion, 0) if champion != "-" else 0

    start_map = _rank_map_until(start_day - 1)
    end_map = _rank_map_until(end_day)
    rising_name, rising_delta = None, -10**9
    falling_name, falling_delta = None, 10**9
    for name in PARTICIPANTS:
        old_rank = start_map.get(name, 999)
        new_rank = end_map.get(name, 999)
        if old_rank == 999 or new_rank == 999:
            continue
        delta = old_rank - new_rank
        if delta > rising_delta:
            rising_delta, rising_name = delta, name
        if delta < falling_delta:
            falling_delta, falling_name = delta, name

    top_day_name, top_day_score, top_day = _biggest_daily_score(start_day, end_day)
    legends_map = stats_range.get("daily_wins", {})
    top_legend_name = "-"
    top_legend_count = 0
    if legends_map:
        top_legend_name, top_legend_count = max(legends_map.items(), key=lambda x: x[1])

    width = 1400
    top5 = totals.most_common(5)
    height = max(1080, 860 + len(top5) * 62)
    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, f"تقرير الفترة من اليوم {start_day} إلى {end_day}", "فانتزي المصيف 2026", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=230, bottom_pad=110, accent="#F59E0B")

    cards = [
        (80, fy1+15, 430, fy1+160, "بطل الفترة", f"{_clean_display_name(champion)}\n{champ_points} نقطة"),
        (525, fy1+15, 875, fy1+160, "أكثر اللاعبين صعودًا", _movement_text(rising_name, start_map.get(rising_name), end_map.get(rising_name))),
        (970, fy1+15, 1320, fy1+160, "أكثر اللاعبين تراجعًا", _movement_text(falling_name, start_map.get(falling_name), end_map.get(falling_name))),
        (80, fy1+205, 652, fy1+355, "أعلى نقاط يومية في الفترة", f"{_clean_display_name(top_day_name)}\n{top_day_score} نقطة — الجولة {top_day}"),
        (748, fy1+205, 1320, fy1+355, "أكثر من فاز بأسطورة اليوم", f"{_clean_display_name(top_legend_name)}\n{top_legend_count} مرات"),
    ]
    for x1,y1,x2,y2,t,v in cards:
        rounded_rect(draw,(x1,y1,x2,y2), radius=30, fill="#091122DD", outline="#FFFFFF25", width=2)
        draw_text(draw, ((x1+x2)//2, y1+38), t, get_font(27), fill="#E5E7EB")
        draw_text(draw, ((x1+x2)//2, y1+95), v, get_font(35), fill="#FFFFFF", max_width=x2-x1-30)

    y = fy1 + 405
    draw_text(draw, (700, y), "أفضل 5 في الفترة", get_font(36), fill="#FDE68A")
    y += 55
    for i, (pname, pts) in enumerate(top5, start=1):
        accent = "#F59E0B" if i == 1 else v16_accent(i)
        rounded_rect(draw,(170,y,1230,y+55), radius=18, fill="#0B1020", outline=accent, width=2)
        draw_text(draw,(1110,y+28), f"{i}. {_clean_display_name(pname)}", get_font(29), fill="#FFFFFF", max_width=650)
        draw_text(draw,(300,y+28), f"{pts} نقطة", get_font(29), fill="#FDE68A")
        y += 65

    footer_event(draw, width, height)
    path = os.path.join(GENERATED_DIR, f"period_report_{start_day}_{end_day}.png")
    img.save(path, quality=95)
    return path



# ============================================================
# V26 FINAL OVERRIDES
# اعتماد التعديلات النهائية:
# - ستايل Games of the Day شبه المرجع بخلفية ثابتة
# - كابتشن نظيف للتصاميم بدون نموذج فانتزي
# - إصلاح صورة اليوم، بطاقة المشارك، سجل الأساطير، تقرير الفترة
# - إيقاف صور الإحصائيات غير المقروءة كجداول
# ============================================================

def paste_event_logo(img, width, y=48):
    # حسب الاعتماد: لا نضع شعار استراحة المصيف على تصاميم الفانتزي
    return

def draw_design_header(draw, width, title, subtitle, img=None):
    draw_text(draw, (width//2, 96), title, get_font(54), fill="#FFFFFF", max_width=width-120)
    draw_text(draw, (width//2, 160), subtitle, get_font(34), fill="#FDE68A", max_width=width-160)
    draw.line((210, 213, width-210, 213), fill="#FFFFFF45", width=1)

def _clean_display_name(value):
    s = str(value or "").strip()
    # إزالة رموز قد تظهر كمربعات داخل الصور
    s = re.sub(r"[\U00010000-\U0010ffff]", "", s)
    s = s.replace("🥇","").replace("🥈","").replace("🥉","")
    s = s.replace("🐐","").replace("👑","").replace("✅","")
    s = s.replace("—", "-").replace("–", "-").replace("|", "-")
    return re.sub(r"\s+", " ", s).strip()

def _safe_rank_label(idx):
    return str(idx)

def build_design_matches_caption(day_name, matches):
    lines = ["🏆 مونديال المصيف 2026 🏆", f"🔥 مباريات اليوم ( {day_name} ) 🔥", ""]
    for a, b, t in matches:
        lines.append(f"{a} × {b} — {t}")
    lines.append("")
    lines.append("المصيف ينقل لكم الحدث")
    return "\n".join(lines)

def build_design_results_caption(day_name, results):
    lines = ["🏆 نتائج مباريات اليوم 🏆", f"اليوم ( {day_name} )", ""]
    for a, sa, sb, b in results:
        lines.append(f"{a} {sa} - {sb} {b}")
    lines.append("")
    lines.append("المصيف ينقل لكم الحدث")
    return "\n".join(lines)

def _games_day_background(width, height):
    img = Image.new("RGB", (width, height), "#06152F")
    draw = ImageDraw.Draw(img)

    # تدرج ثابت أزرق
    for y in range(height):
        t = y / max(1, height)
        r = int(2 + 5*t)
        g = int(14 + 20*t)
        b = int(45 + 35*t)
        draw.line((0, y, width, y), fill=(r, g, b))

    overlay = Image.new("RGBA", (width, height), (0,0,0,0))
    od = ImageDraw.Draw(overlay)

    # دوائر/أشكال الخلفية الثابتة
    od.ellipse((160, -185, width-150, 490), fill=(15, 92, 190, 112))
    od.ellipse((width*0.62, 395, width*1.18, 1010), fill=(0, 102, 230, 50))
    od.ellipse((width*0.68, 720, width*1.12, 1240), fill=(0, 95, 215, 42))
    od.rectangle((width*0.63, 430, width*0.90, 640), fill=(0, 115, 255, 36))
    od.rectangle((width*0.90, 520, width*1.08, 790), fill=(0, 105, 230, 42))
    od.ellipse((width-80, 360, width+250, 720), fill=(0, 115, 255, 55))

    # ظل تمثال/شكل جانبي يسار بروح المرجع، مو شعار رسمي
    stat = Image.new("RGBA", (width, height), (0,0,0,0))
    sd = ImageDraw.Draw(stat)
    sx = int(width*0.08)
    base_y = int(height*0.80)
    blue = (20, 95, 190, 95)
    sd.polygon([(sx-130, base_y+180),(sx+210, base_y+180),(sx+135, base_y-250),(sx-40, base_y-290)], fill=blue)
    sd.ellipse((sx-45, base_y-420, sx+75, base_y-300), fill=(20,95,190,82))
    # تاج بسيط
    for ang in [-70,-45,-20,10,35,60]:
        x2 = sx + 15 + int(115 * (ang/70))
        sd.line((sx+15, base_y-405, x2, base_y-540), fill=(20,95,190,75), width=9)
    # اليد/الشعلة
    sd.line((sx-125, base_y-350, sx-205, base_y-625), fill=(20,95,190,85), width=34)
    sd.ellipse((sx-245, base_y-700, sx-165, base_y-615), fill=(20,95,190,90))
    sd.polygon([(sx-205, base_y-735),(sx-235, base_y-665),(sx-170, base_y-670)], fill=(20,95,190,85))
    stat = stat.filter(ImageFilter.GaussianBlur(2))
    overlay = Image.alpha_composite(overlay, stat)

    overlay = overlay.filter(ImageFilter.GaussianBlur(8))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)

    # خطوط خفيفة علوية
    for i in range(0, 70, 7):
        draw.arc((80+i*2, 20+i, width-80-i*2, 720+i), 190, 335, fill="#1D6FEA45", width=2)

    return img, draw

def _draw_games_footer(draw, img, width, height):
    # الراعي الأول: MASEEF SPORTS
    pill_y = height - 142
    pill_w, pill_h = 225, 58
    x1 = width//2 - pill_w - 25
    x2 = width//2 + 25

    for x, label1, label2 in [
        (x1, "MASEEF", "SPORTS"),
        (x2, "MASEEF", "2026"),
    ]:
        rounded_rect(draw, (x, pill_y, x+pill_w, pill_y+pill_h), radius=18, fill="#071633BB", outline="#FFFFFFAA", width=2)
        draw_text(draw, (x+pill_w//2, pill_y+22), label1, get_font(22), fill="#FFFFFF", max_width=pill_w-35)
        draw_text(draw, (x+pill_w//2, pill_y+44), label2, get_font(16), fill="#E5E7EB", max_width=pill_w-35)

    # بادج السعودية يمين
    bx = width - 150
    by = height - 225
    draw_text(draw, (bx+55, by-28), "TIME OF DAY", get_font(18), fill="#FFFFFF")
    rounded_rect(draw, (bx, by, bx+110, by+70), radius=20, fill="#087A36", outline="#A7F3D0", width=1)
    draw_text(draw, (bx+55, by+35), "SA", get_font(30), fill="#FFFFFF")
    rounded_rect(draw, (bx+8, by+76, bx+102, by+128), radius=18, fill="#0B7F36", outline="#A7F3D0", width=1)
    draw_text(draw, (bx+55, by+102), "KSA", get_font(22), fill="#FFFFFF")
    draw_text(draw, (bx+55, by+155), "SAUDI ARABIA", get_font(17), fill="#FFFFFF")

    # النص السفلي
    draw.line((110, height-55, width//2-240, height-55), fill="#D4AF37", width=2)
    draw.line((width//2+240, height-55, width-110, height-55), fill="#D4AF37", width=2)
    draw_text(draw, (width//2, height-55), "المصيف ينقل لكم الحدث", get_font(32), fill="#FBBF24")

def _draw_games_header(draw, width, title, date_text):
    draw_text(draw, (width//2, 70), "MONDIAL AL MASEEF 2026", get_font(34), fill="#FFFFFF")
    draw_text(draw, (width//2, 165), title, get_font(78), fill="#FFFFFF", max_width=width-80)
    rounded_rect(draw, (width//2-230, 255, width//2+230, 312), radius=18, fill="#FBBF24", outline="#00000070", width=2)
    draw_text(draw, (width//2, 285), date_text, get_font(28), fill="#061633", max_width=430)

def create_matches_style2_image(day_name, matches):
    ensure_generated_dir()
    count = max(len(matches), 1)
    width = 1200
    height = 1500 if count >= 4 else 1280
    img, draw = _games_day_background(width, height)

    _draw_games_header(draw, width, "GAMES OF THE DAY", f"اليوم {day_name}")

    if count == 1:
        card_h, gap, y = 245, 0, 490
    elif count == 2:
        card_h, gap, y = 205, 34, 430
    elif count == 3:
        card_h, gap, y = 180, 24, 400
    elif count == 4:
        card_h, gap, y = 155, 22, 400
    else:
        card_h, gap, y = 132, 17, 375

    x1, x2 = 270, 890
    for a, b, t in matches[:6]:
        rounded_rect(draw, (x1, y, x2, y+card_h), radius=26, fill="#0638A5", outline="#13B6EA", width=3)
        draw.line((x1+8, y+card_h-3, x2-8, y+card_h-3), fill="#EF4444", width=4)
        draw.line((x1+80, y+4, x2-80, y+4), fill="#22C55E", width=3)
        draw.line((x1+260, y+4, x1+315, y+4), fill="#FBBF24", width=3)
        draw.arc((x1, y, x1+58, y+58), 180, 270, fill="#14B8F5", width=7)
        draw.arc((x2-58, y+card_h-58, x2, y+card_h), 0, 90, fill="#EF4444", width=7)

        cy = y + card_h//2
        flag_w = min(150, card_h-42)
        paste_flag(img, a, (x1+36, cy-flag_w//2, x1+36+flag_w, cy+flag_w//2))
        paste_flag(img, b, (x2-36-flag_w, cy-flag_w//2, x2-36, cy+flag_w//2))

        name_font = get_font(25 if count >= 4 else 29)
        draw_text(draw, (x1+36+flag_w//2, y+card_h-28), _clean_display_name(a), name_font, fill="#FFFFFF", max_width=210)
        draw_text(draw, (x2-36-flag_w//2, y+card_h-28), _clean_display_name(b), name_font, fill="#FFFFFF", max_width=210)

        tm, period = _ampm_from_time(t)
        draw_text(draw, (width//2, cy-14), tm, get_font(50 if count >= 4 else 62), fill="#FBBF24")
        if period:
            draw_text(draw, (width//2, cy+39), period, get_font(29 if count >= 4 else 36), fill="#FBBF24")
        y += card_h + gap

    _draw_games_footer(draw, img, width, height)
    path = os.path.join(GENERATED_DIR, f"matches_style2_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path

def create_match_results_style2_image(day_name, results):
    ensure_generated_dir()
    count = max(len(results), 1)
    width = 1200
    height = 1500 if count >= 4 else 1280
    img, draw = _games_day_background(width, height)

    _draw_games_header(draw, width, "MATCH RESULTS", f"اليوم {day_name}")

    if count == 1:
        card_h, gap, y = 245, 0, 490
    elif count == 2:
        card_h, gap, y = 205, 34, 430
    elif count == 3:
        card_h, gap, y = 180, 24, 400
    elif count == 4:
        card_h, gap, y = 155, 22, 400
    else:
        card_h, gap, y = 132, 17, 375

    x1, x2 = 270, 890
    for a, sa, sb, b in results[:6]:
        rounded_rect(draw, (x1, y, x2, y+card_h), radius=26, fill="#0638A5", outline="#13B6EA", width=3)
        draw.line((x1+8, y+card_h-3, x2-8, y+card_h-3), fill="#EF4444", width=4)
        draw.line((x1+80, y+4, x2-80, y+4), fill="#22C55E", width=3)
        cy = y + card_h//2
        flag_w = min(145, card_h-46)
        paste_flag(img, a, (x1+36, cy-flag_w//2, x1+36+flag_w, cy+flag_w//2))
        paste_flag(img, b, (x2-36-flag_w, cy-flag_w//2, x2-36, cy+flag_w//2))

        name_font = get_font(25 if count >= 4 else 29)
        draw_text(draw, (x1+36+flag_w//2, y+card_h-28), _clean_display_name(a), name_font, fill="#FFFFFF", max_width=210)
        draw_text(draw, (x2-36-flag_w//2, y+card_h-28), _clean_display_name(b), name_font, fill="#FFFFFF", max_width=210)

        rounded_rect(draw, (width//2-96, cy-42, width//2+96, cy+42), radius=18, fill="#FBBF24", outline="#00000088", width=2)
        draw_text(draw, (width//2, cy), f"{sa} - {sb}", get_font(48 if count >= 4 else 58), fill="#061633")
        y += card_h + gap

    _draw_games_footer(draw, img, width, height)
    path = os.path.join(GENERATED_DIR, f"results_style2_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path

def create_scorers_style2_image(items):
    ensure_generated_dir()
    items = sorted(items, key=lambda x: (-x[1], x[0]))[:12]
    count = max(len(items), 1)
    width = 1200
    height = 1500
    img, draw = _games_day_background(width, height)
    _draw_games_header(draw, width, "TOP SCORERS", "هدافين البطولة")

    if count <= 3:
        row_h, gap, y = 110, 24, 430
    else:
        row_h, gap, y = 80, 14, 405

    x1, x2 = 230, 950
    for i, (name, goals, team) in enumerate(items, start=1):
        rounded_rect(draw, (x1, y, x2, y+row_h), radius=24, fill="#0638A5", outline=v16_accent(i), width=3)
        cy = y + row_h//2
        draw_text(draw, (x2-55, cy), str(i), get_font(34), fill="#FBBF24")
        if team:
            paste_flag(img, team, (x2-145, cy-30, x2-85, cy+30))
        draw_text(draw, (width//2, cy), _clean_display_name(name), get_font(30 if count <= 3 else 24), fill="#FFFFFF", max_width=420)
        draw_text(draw, (x1+115, cy), f"{goals} أهداف", get_font(27 if count <= 3 else 23), fill="#FBBF24")
        y += row_h + gap

    _draw_games_footer(draw, img, width, height)
    path = os.path.join(GENERATED_DIR, "scorers_style2.png")
    img.save(path, quality=95)
    return path

def create_group_style2_image(group_title, rows):
    ensure_generated_dir()
    width, height = 1200, 1500
    img, draw = _games_day_background(width, height)
    title = clean_group_title_for_design(group_title)
    if not str(title).startswith("المجموعة"):
        title = f"المجموعة {title}"
    _draw_games_header(draw, width, "GROUP STANDINGS", title)

    x1, x2 = 210, 990
    y = 435
    row_h, gap = 105, 18
    for i, (team, played, diff, pts) in enumerate(rows[:6], start=1):
        rounded_rect(draw, (x1, y, x2, y+row_h), radius=24, fill="#0638A5", outline=v16_accent(i), width=3)
        cy = y + row_h//2
        draw_text(draw, (x2-52, cy), str(i), get_font(32), fill="#FBBF24")
        paste_flag(img, team, (x2-145, cy-33, x2-80, cy+33))
        draw_text(draw, (x2-300, cy), _clean_display_name(team), get_font(30), fill="#FFFFFF", max_width=260)
        draw_text(draw, (x1+355, cy), f"لعب {played}", get_font(24), fill="#E5E7EB")
        draw_text(draw, (x1+230, cy), f"{int(diff):+d}", get_font(28), fill="#E5E7EB")
        draw_text(draw, (x1+90, cy), str(pts), get_font(34), fill="#FBBF24")
        y += row_h + gap

    _draw_games_footer(draw, img, width, height)
    path = os.path.join(GENERATED_DIR, f"group_style2_{_safe_filename(group_title)}.png")
    img.save(path, quality=95)
    return path

def create_all_groups_image(groups):
    ensure_generated_dir()
    width, height = 1800, 2400
    img, draw = _games_day_background(width, height)
    draw_text(draw, (width//2, 90), "MONDIAL AL MASEEF 2026", get_font(40), fill="#FFFFFF")
    draw_text(draw, (width//2, 170), "ALL GROUP STANDINGS", get_font(72), fill="#FFFFFF", max_width=width-160)
    draw_text(draw, (width//2, 235), "ترتيب جميع المجموعات", get_font(36), fill="#FBBF24")

    cols = 3
    margin_x, gap_x = 75, 35
    card_w = (width - 2*margin_x - (cols-1)*gap_x) // cols
    card_h = 470
    start_y = 320
    gap_y = 38

    for idx, (title, rows) in enumerate(groups[:12]):
        c = idx % cols
        r = idx // cols
        x = margin_x + c*(card_w+gap_x)
        y = start_y + r*(card_h+gap_y)
        rounded_rect(draw, (x, y, x+card_w, y+card_h), radius=28, fill="#0638A5EE", outline="#14B8F5", width=3)
        rounded_rect(draw, (x+18, y+18, x+card_w-18, y+68), radius=18, fill="#FBBF24", outline="#00000055", width=1)
        gt = clean_group_title_for_design(title)
        if not str(gt).startswith("المجموعة"):
            gt = f"المجموعة {gt}"
        draw_text(draw, (x+card_w//2, y+43), gt, get_font(26), fill="#061633", max_width=card_w-40)
        yy = y + 92
        for pos, (team, played, diff, pts) in enumerate(rows[:4], start=1):
            rounded_rect(draw, (x+20, yy, x+card_w-20, yy+72), radius=16, fill="#061633AA", outline="#FFFFFF22", width=1)
            cy = yy + 36
            draw_text(draw, (x+card_w-44, cy), str(pos), get_font(22), fill="#FBBF24")
            paste_flag(img, team, (x+card_w-125, cy-24, x+card_w-75, cy+24))
            draw_text(draw, (x+card_w-245, cy), _clean_display_name(team), get_font(21), fill="#FFFFFF", max_width=185)
            draw_text(draw, (x+155, cy), str(pts), get_font(26), fill="#FBBF24")
            draw_text(draw, (x+75, cy), f"{int(diff):+d}", get_font(21), fill="#E5E7EB")
            yy += 84

    draw_text(draw, (width//2, height-70), "المصيف ينقل لكم الحدث", get_font(36), fill="#FBBF24")
    path = os.path.join(GENERATED_DIR, "all_groups_style2.png")
    img.save(path, quality=95)
    return path

def create_daily_result_image(day, goals_count=None, clean_sheets=None):
    ensure_generated_dir()
    data = build_day_summary(day)
    rows = data["sorted_rows"]
    matchups = generate_matchups_for_day(day)
    pairs = matchups.get("pairs", [])
    match_rows = len(pairs) + (1 if matchups.get("bye") else 0)

    goals_lines = [f"{_clean_display_name(p)} - {c}" for p, c in (goals_count or {}).items()] if goals_count else [f"{_clean_display_name(p)} - {pts} نقطة" for p, pts in data["player_points"].most_common(5)]
    clean_lines = [_clean_display_name(x) for x in (clean_sheets or list(data["keeper_points"].keys()) or ["لا يوجد"])]
    goals_lines = goals_lines or ["لا يوجد"]
    clean_lines = clean_lines or ["لا يوجد"]

    width = 1400
    row_h = 44
    match_box_h = max(165, 85 + match_rows * 36)
    height = max(1580, 520 + len(rows)*50 + match_box_h + 300)
    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, f"فانتزي المصيف 2026 - اليوم {ordinal_day(day)}", "أسطورة اليوم وترتيب المشاركين", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=235, bottom_pad=95, accent="#8B5CF6")

    legends = " + ".join(_clean_display_name(w) for w in data["winners"]) if data["winners"] else "لا يوجد"
    rounded_rect(draw, (90, fy1 + 18, 1010, fy1 + 145), radius=32, fill="#7C3AEDDD", outline="#FFFFFF40", width=2)
    draw_text(draw, (550, fy1 + 55), "أسطورة اليوم", get_font(27), fill="#FFFFFF")
    draw_text(draw, (550, fy1 + 106), f"{legends} - {data['max_score']} نقطة", get_font(39), fill="#FFF6D6", max_width=850)
    rounded_rect(draw, (1040, fy1 + 18, 1310, fy1 + 145), radius=26, fill="#F59E0BDD", outline="#FFFFFF33", width=2)
    draw_text(draw, (1175, fy1 + 55), "المشاركون", get_font(24), fill="#FFFFFF")
    draw_text(draw, (1175, fy1 + 106), f"{len(data['participants'])}", get_font(44), fill="#FFFFFF")

    y = fy1 + 175
    rounded_rect(draw, (90, y, 1310, y + 50), radius=16, fill="#05070D", outline="#FFFFFF55", width=1)
    headers = [(1180, "المشارك"), (870, "النقاط"), (650, "الحارس"), (430, "الكابتن"), (190, "المركز")]
    for x, t in headers:
        draw_text(draw, (x, y + 25), t, get_font(21), fill="#FDE68A")
    y += 60

    for idx, r in enumerate(rows, start=1):
        accent = "#F59E0B" if idx == 1 else "#FFFFFF22"
        fill = "#1A1407" if idx == 1 else ("#0B1020" if idx % 2 else "#10172A")
        rounded_rect(draw, (90, y, 1310, y + row_h), radius=14, fill=fill, outline=accent, width=2 if idx == 1 else 1)
        cy = y + row_h // 2
        draw_text(draw, (1180, cy), _clean_display_name(r["participant"]), get_font(22), fill="#FFFFFF", max_width=300)
        draw_text(draw, (870, cy), str(r["total"]), get_font(25), fill="#FDE68A")
        draw_text(draw, (650, cy), f"+{r['keeper_points']}", get_font(22), fill="#A7F3D0")
        draw_text(draw, (430, cy), f"+{r['captain_points']}", get_font(22), fill="#C4B5FD")
        draw_text(draw, (190, cy), str(idx), get_font(23), fill="#FFFFFF")
        y += row_h + 6

    rounded_rect(draw, (90, y + 18, 1310, y + 18 + match_box_h), radius=26, fill="#091122DD", outline="#22C55E", width=2)
    draw_text(draw, (700, y + 52), "مواجهات اليوم", get_font(27), fill="#FDE68A")
    yy = y + 88
    if not pairs and not matchups.get("bye"):
        draw_text(draw, (700, yy), "لا توجد مواجهات", get_font(22), fill="#FFFFFF")
    else:
        for p in pairs:
            line = f"{_clean_display_name(p['a'])} {p['a_points']} - {p['b_points']} {_clean_display_name(p['b'])} - الفائز: {_clean_display_name(p['winner'])}"
            draw_text(draw, (700, yy), line, get_font(20), fill="#FFFFFF", max_width=1120)
            yy += 34
        if matchups.get("bye"):
            draw_text(draw, (700, yy), f"راحة الجولة: {_clean_display_name(matchups['bye'])}", get_font(20), fill="#FDE68A")

    y = y + 18 + match_box_h + 26
    box_h = 205
    rounded_rect(draw, (90, y, 675, y + box_h), radius=26, fill="#091122DD", outline="#A855F7", width=2)
    rounded_rect(draw, (725, y, 1310, y + box_h), radius=26, fill="#091122DD", outline="#06B6D4", width=2)
    draw_text(draw, (382, y + 38), "الهدافين", get_font(30), fill="#FFFFFF")
    draw_text(draw, (382, y + 120), "\n".join(goals_lines[:5]), get_font(23), fill="#E5E7EB", max_width=500)
    draw_text(draw, (1018, y + 38), "الكلين شيت", get_font(30), fill="#FFFFFF")
    draw_text(draw, (1018, y + 120), "\n".join(clean_lines[:5]), get_font(23), fill="#E5E7EB", max_width=500)

    draw_text(draw, (width//2, height-48), "المصيف ينقل لكم الحدث", get_font(26), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, f"daily_result_day_{day}.png")
    img.save(path, quality=95)
    return path

def create_legends_image(start_day=1, end_day=31):
    ensure_generated_dir()
    stats = collect_stats(start_day, end_day)
    days = stats["days"]
    width = 1200
    count = max(len(days), 1)
    row_h = 82
    gap = 12
    height = max(820, 265 + count*(row_h+gap) + 125)
    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, "سجل أساطير الفانتزي", f"من اليوم {start_day} إلى اليوم {end_day}", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=220, bottom_pad=100, accent="#8B5CF6")

    y = fy1 + 24
    rounded_rect(draw, (92, y, width-92, y+56), radius=18, fill="#05070D", outline="#FFFFFF55", width=1)
    draw_text(draw, (1010, y+28), "اليوم", get_font(22), fill="#FDE68A")
    draw_text(draw, (650, y+28), "أسطورة اليوم", get_font(22), fill="#FDE68A")
    draw_text(draw, (220, y+28), "النقاط", get_font(22), fill="#FDE68A")
    y += 72

    if not days:
        draw_text(draw, (width//2, y+35), "لا توجد أيام في هذا النطاق", get_font(24), fill="#FFFFFF")
    else:
        for idx, day in enumerate(days, start=1):
            info = stats["per_day"].get(day, {})
            winners = " + ".join(_clean_display_name(w) for w in info.get("winners", [])) if info.get("winners") else "لا يوجد"
            max_score = info.get("max_score", 0)
            accent = "#F59E0B" if idx == 1 else v16_accent(idx)
            rounded_rect(draw, (92, y, width-92, y+row_h), radius=20, fill="#0B1020", outline=accent, width=2)
            cy = y + row_h//2
            draw_text(draw, (1010, cy), f"اليوم {ordinal_day(day)}", get_font(23), fill="#FFFFFF")
            draw_text(draw, (650, cy), winners, get_font(22), fill="#FDE68A", max_width=500)
            draw_text(draw, (220, cy), f"{max_score} نقطة", get_font(23), fill="#A7F3D0")
            y += row_h + gap

    draw_text(draw, (width//2, height-44), "المصيف ينقل لكم الحدث", get_font(24), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, f"legends_{start_day}_{end_day}.png")
    img.save(path, quality=95)
    return path

def create_participant_card_image(name, start_day=1, end_day=31):
    ensure_generated_dir()
    stats = collect_stats(start_day, end_day)
    if name not in PARTICIPANTS:
        raise ValueError("اسم المشارك غير موجود")
    rank = stats["ranking"].index(name) + 1 if name in stats["ranking"] else "-"
    total = stats["totals"].get(name, 0)
    day_scores = stats["scores_by_day"].get(name, {})
    best_day = max(day_scores, key=lambda d: day_scores[d]) if day_scores else "-"
    worst_day = min(day_scores, key=lambda d: day_scores[d]) if day_scores else "-"
    best_score = day_scores.get(best_day, 0) if best_day != "-" else 0
    worst_score = day_scores.get(worst_day, 0) if worst_day != "-" else 0
    best_cap, best_cap_pts, best_cap_day = participant_best_captain(name, start_day, end_day)
    wins = matchup_wins_map(start_day, end_day).get(name, 0)
    pc = stats["participation_count"].get(name, 0)
    days_count = len(stats["days"])
    pct = f"{round(pc / days_count * 100, 1)}%" if days_count else "0%"

    width, height = 1200, 940
    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, "بطاقة مشارك فانتزي المصيف", _clean_display_name(name), img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=225, bottom_pad=105, accent="#8B5CF6")

    best_day_text = f"اليوم {best_day}\n{best_score} نقاط" if best_day != "-" else "-"
    worst_day_text = f"اليوم {worst_day}\n{worst_score} نقاط" if worst_day != "-" else "-"
    cap_text = f"{_clean_display_name(best_cap)}\n{best_cap_pts} نقاط" if best_cap else "-"

    cards = [
        (90, fy1+20, 360, fy1+145, "المركز", f"#{rank}"),
        (465, fy1+20, 735, fy1+145, "النقاط", str(total)),
        (840, fy1+20, 1110, fy1+145, "أسطورة اليوم", str(stats["daily_wins"].get(name,0))),
        (90, fy1+185, 360, fy1+315, "أفضل يوم", best_day_text),
        (465, fy1+185, 735, fy1+315, "أسوأ يوم", worst_day_text),
        (840, fy1+185, 1110, fy1+315, "نسبة المشاركة", pct),
        (90, fy1+355, 545, fy1+500, "أفضل كابتن", cap_text),
        (655, fy1+355, 1110, fy1+500, "انتصارات المواجهات اليومية", str(wins)),
    ]
    for x1,y1,x2,y2,t,v in cards:
        rounded_rect(draw,(x1,y1,x2,y2), radius=28, fill="#091122DD", outline="#FFFFFF25", width=2)
        draw_text(draw, ((x1+x2)//2, y1+34), t, get_font(25), fill="#E5E7EB", max_width=x2-x1-20)
        draw_text(draw, ((x1+x2)//2, y1+92), v, get_font(31), fill="#FFFFFF", max_width=x2-x1-30)

    draw_text(draw, (width//2, height-44), "المصيف ينقل لكم الحدث", get_font(24), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, f"participant_card_{_safe_filename(name)}.png")
    img.save(path, quality=95)
    return path

def _period_start_rank_map(start_day):
    before = _rank_map_until(start_day - 1)
    if before:
        return before
    return _rank_map_until(start_day)

def _movement_label(name, old_rank, new_rank):
    if not name or old_rank is None or new_rank is None:
        return "لا يوجد تغير"
    return f"{_clean_display_name(name)}\nمن المركز {old_rank} إلى {new_rank}"

def create_period_report_image(start_day, end_day):
    ensure_generated_dir()
    days, totals, cap_points, keeper_points, player_impact, player_zero, active_counts = period_stats(start_day, end_day)
    if not days:
        raise ValueError("ما فيه أيام في الفترة")

    stats_range = collect_stats(start_day, end_day)
    champion = totals.most_common(1)[0][0] if totals else "-"
    champ_points = totals.get(champion, 0) if champion != "-" else 0

    start_map = _period_start_rank_map(start_day)
    end_map = _rank_map_until(end_day)
    rising_name = falling_name = None
    rising_delta = 0
    falling_delta = 0
    for name in PARTICIPANTS:
        if name not in start_map or name not in end_map:
            continue
        delta = start_map[name] - end_map[name]
        if delta > rising_delta:
            rising_delta, rising_name = delta, name
        if delta < falling_delta:
            falling_delta, falling_name = delta, name

    rising_text = _movement_label(rising_name, start_map.get(rising_name), end_map.get(rising_name)) if rising_name else "لا يوجد تغير"
    falling_text = _movement_label(falling_name, start_map.get(falling_name), end_map.get(falling_name)) if falling_name else "لا يوجد تغير"

    top_day_name, top_day_score, top_day = _biggest_daily_score(start_day, end_day)
    legends_map = stats_range.get("daily_wins", {})
    if legends_map:
        top_legend_name, top_legend_count = max(legends_map.items(), key=lambda x: x[1])
    else:
        top_legend_name, top_legend_count = "-", 0

    width = 1400
    top5 = totals.most_common(5)
    height = max(1080, 865 + len(top5) * 62)
    img, draw = design_canvas(None, width, height, "purple")
    draw_design_header(draw, width, f"تقرير الفترة من اليوم {start_day} إلى {end_day}", "فانتزي المصيف 2026", img)
    fx1, fy1, fx2, fy2 = draw_broadcast_inner_frame(draw, width, height, top=230, bottom_pad=105, accent="#F59E0B")

    cards = [
        (80, fy1+15, 430, fy1+160, "بطل الفترة", f"{_clean_display_name(champion)}\n{champ_points} نقطة"),
        (525, fy1+15, 875, fy1+160, "أكثر اللاعبين صعودًا", rising_text),
        (970, fy1+15, 1320, fy1+160, "أكثر اللاعبين تراجعًا", falling_text),
        (80, fy1+205, 652, fy1+355, "أعلى نقاط يومية في الفترة", f"{_clean_display_name(top_day_name)}\n{top_day_score} نقطة\nالجولة {top_day}"),
        (748, fy1+205, 1320, fy1+355, "أكثر من فاز بأسطورة اليوم", f"{_clean_display_name(top_legend_name)}\n{top_legend_count} مرات"),
    ]
    for x1,y1,x2,y2,t,v in cards:
        rounded_rect(draw,(x1,y1,x2,y2), radius=30, fill="#091122DD", outline="#FFFFFF25", width=2)
        draw_text(draw, ((x1+x2)//2, y1+38), t, get_font(25), fill="#E5E7EB", max_width=x2-x1-20)
        draw_text(draw, ((x1+x2)//2, y1+101), v, get_font(31), fill="#FFFFFF", max_width=x2-x1-30)

    y = fy1 + 405
    draw_text(draw, (700, y), "أفضل 5 في الفترة", get_font(36), fill="#FDE68A")
    y += 55
    for i, (pname, pts) in enumerate(top5, start=1):
        accent = "#F59E0B" if i == 1 else v16_accent(i)
        rounded_rect(draw,(170,y,1230,y+55), radius=18, fill="#0B1020", outline=accent, width=2)
        draw_text(draw,(1110,y+28), f"{i}. {_clean_display_name(pname)}", get_font(28), fill="#FFFFFF", max_width=650)
        draw_text(draw,(300,y+28), f"{pts} نقطة", get_font(28), fill="#FDE68A")
        y += 65

    draw_text(draw, (width//2, height-45), "المصيف ينقل لكم الحدث", get_font(24), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, f"period_report_{start_day}_{end_day}.png")
    img.save(path, quality=95)
    return path

async def dashboard_sheet_image_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        args = update.message.text.split(maxsplit=3)
        start_day = int(args[1]) if len(args) >= 2 and args[1].isdigit() else 1
        end_day = int(args[2]) if len(args) >= 3 and args[2].isdigit() else max(get_existing_days(1, 31) or [1])
        sheet = args[3].strip() if len(args) >= 4 else ""
        if "ترتيب" in sheet:
            path = create_overall_ranking_image(start_day, end_day)
            await send_photo_path(update, path, f"✅ الترتيب العام من اليوم {start_day} إلى {end_day}")
        elif "أساطير" in sheet or "اساطير" in sheet or "سجل" in sheet:
            path = create_legends_image(start_day, end_day)
            await send_photo_path(update, path, f"✅ سجل الأساطير من اليوم {start_day} إلى {end_day}")
        else:
            await update.message.reply_text(
                "صور الإحصائيات كجداول Excel غير معتمدة لأنها غير مقروءة داخل تيليجرام.\n"
                "استخدم /احصائيات للملف، أو /صور_الاحصائيات للصور المعتمدة."
            )
    except Exception as e:
        await update.message.reply_text(f"تعذر تجهيز صورة الإحصائيات ❌\n{e}")

async def all_dashboard_images_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        m = re.search(r"(\d+)\s+(\d+)", update.message.text)
        if m:
            start_day, end_day = int(m.group(1)), int(m.group(2))
        else:
            start_day, end_day = 1, max(get_existing_days(1, 31) or [1])
        await update.message.reply_text("جاري تجهيز الصور المعتمدة فقط ✅")
        await send_photo_path(update, create_overall_ranking_image(start_day, end_day), f"✅ الترتيب العام من اليوم {start_day} إلى {end_day}")
        await send_photo_path(update, create_legends_image(start_day, end_day), f"✅ سجل الأساطير من اليوم {start_day} إلى {end_day}")
        await send_photo_path(update, create_period_report_image(start_day, end_day), f"✅ تقرير الفترة {start_day} - {end_day}")
        await update.message.reply_text("تم إيقاف صور الجداول الصغيرة غير المقروءة. ملف Excel موجود عبر /احصائيات ✅")
    except Exception as e:
        await update.message.reply_text(f"تعذر تجهيز صور الإحصائيات ❌\n{e}")

async def design_matches_style2_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        day_name, matches = parse_matches_text(update.message.text)
        if not matches:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_مباريات_ستايل2\nالسابع\nالبرتغال|الكونغو الديمقراطية|8:00 م")
            return
        path = create_matches_style2_image(day_name, matches)
        await send_photo_path(update, path, build_design_matches_caption(day_name, matches))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم مباريات ستايل2 ❌\n{e}")

async def design_match_results_style2_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        day_name, results = parse_match_results_design_text(update.message.text)
        if not results:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_نتائج_مباريات_ستايل2\nالسابع\nالبرتغال|2|1|الكونغو الديمقراطية")
            return
        path = create_match_results_style2_image(day_name, results)
        await send_photo_path(update, path, build_design_results_caption(day_name, results))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم نتائج ستايل2 ❌\n{e}")

async def design_scorers_style2_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        items = parse_scorers_text(update.message.text)
        if not items:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_هدافين_ستايل2\nراؤول خيمينيز|4|المكسيك")
            return
        path = create_scorers_style2_image(items)
        await send_photo_path(update, path, build_top_scorers_caption(items))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم الهدافين ستايل2 ❌\n{e}")

async def design_group_style2_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        group_title, rows = parse_group_standing_text(update.message.text)
        if not rows:
            await update.message.reply_text("اكتبها كذا:\n/تصميم_ترتيب_مجموعة_ستايل2\nالمجموعة C\nالبرازيل|1|0|1")
            return
        path = create_group_style2_image(group_title, rows)
        await send_photo_path(update, path, build_group_standing_caption(group_title, rows))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم ترتيب المجموعة ستايل2 ❌\n{e}")

async def design_matches_template_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day_name, matches = parse_matches_text(update.message.text)
    if not matches:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_مباريات\nالسادس\nفرنسا|البرازيل|10:00 م\nالأرجنتين|ألمانيا|12:00 ص")
        return
    try:
        path = create_matches_template_image(day_name, matches, use_template=True)
        await send_photo_path(update, path, build_design_matches_caption(day_name, matches))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم المباريات بالقالب ❌\n{e}")

async def design_matches_auto_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    day_name, matches = parse_matches_text(update.message.text)
    if not matches:
        await update.message.reply_text("اكتبها كذا:\n/تصميم_مباريات_تلقائي\nالسادس\nفرنسا|البرازيل|10:00 م")
        return
    try:
        path = create_matches_template_image(day_name, matches, use_template=False)
        await send_photo_path(update, path, build_design_matches_caption(day_name, matches))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم المباريات التلقائي ❌\n{e}")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "الأوامر الأساسية:\n"
        "/اضافه 5\n/نتائج 5\n/اعتماد_نتائج 5\n/احصائيات\n/احصائيات 1 6\n/ترتيب_نص\n\n"
        "أوامر الصور والتقارير:\n"
        "/صورة_اليوم 6\n/صورة_الترتيب 1 6\n/صورة_الاساطير 1 6\n"
        "/صور_الاحصائيات 1 6\n/بطاقة فارس سالم\n/تقرير_الفترة 1 6\n"
        "/تفعيل_الصور_التلقائية\n/إيقاف_الصور_التلقائية\n\n"
        "أوامر التصاميم:\n"
        "/تصميم_مباريات\n/تصميم_مباريات_تلقائي\n"
        "/تصميم_نتائج_مباريات\n/تصميم_نتائج_مباريات_تلقائي\n"
        "/تصميم_ترتيب_مجموعة\n/تصميم_ترتيب_مجموعة_تلقائي\n"
        "/تصميم_هدافين\n/تصميم_هدافين_تلقائي\n\n"
        "التصميم الإضافي المعتمد:\n"
        "/تصميم_مباريات_ستايل2\n/تصميم_نتائج_مباريات_ستايل2\n"
        "/تصميم_هدافين_ستايل2\n/تصميم_ترتيب_مجموعة_ستايل2\n"
        "/تصميم_مباريات_اطار\n/تصميم_نتائج_مباريات_اطار\n"
        "/تصميم_جميع_المجموعات\n\n"
        "أوامر الفحص والنشر:\n"
        "/الأيام\n/فحص 5\n/مشاركين 5\n/اسطورة 5\n/مقارنة 4 5\n/اعلان_اليوم 5\n/ملخص_اليوم 5\n\n"
        "أوامر الأمان:\n"
        "/مسح_نتائج 5\n/مسح_يوم 5\n/مسح_الكل تأكيد\n/استرجاع_آخر\n/قفل_يوم 5\n/فتح_يوم 5\n/معرفي"
    )


# ============================================================
# V27 FINAL — نظام الستايلات 1-4 + الأوامر المختصرة
# اعتماد المستخدم:
# - كل أوامر التصاميم تدعم ستايل 1-4
# - الأوامر المختصرة: /مباريات /انتهت /مجموعة /هدافين /كل_المجموعات
# - النسخ التلقائية تدعم 1-4
# - /نتائج للفانتزي فقط، و/انتهت لتصميم نتائج المباريات
# - دعم صيغ النجمة * بجانب | للصق السريع
# ============================================================

DESIGN_STYLE_KEY = "design_style"
STYLE_NAMES = {
    1: "ستايل 1 - القديم الموجود",
    2: "ستايل 2 - اللوك الجديد الأزرق",
    3: "ستايل 3 - الإطار الأسود",
    4: "ستايل 4 - اللوك الجديد الفخم",
}


def current_design_style():
    try:
        settings = load_settings()
        style = int(settings.get(DESIGN_STYLE_KEY, 4))
        return style if style in (1, 2, 3, 4) else 4
    except Exception:
        return 4


def save_design_style(style):
    style = int(style)
    if style not in (1, 2, 3, 4):
        raise ValueError("الستايل لازم يكون من 1 إلى 4")
    settings = load_settings()
    settings[DESIGN_STYLE_KEY] = style
    save_settings(settings)
    return style


def command_style(text, forced=None):
    if forced in (1, 2, 3, 4):
        return forced
    first = ((text or "").splitlines() or [""])[0].strip()
    parts = first.split()
    if len(parts) >= 2:
        try:
            s = int(parts[1])
            if s in (1, 2, 3, 4):
                return s
        except Exception:
            pass
    return current_design_style()


def _body_lines_after_command(text):
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
    return lines[1:]


def _has_sep(line):
    return any(x in str(line) for x in ["|", "*", "×", " x ", " X "])


def _split_data_parts(line):
    line = normalize_name(line).replace("✕", "×")
    # لا نستبدل حرف x داخل أسماء مثل Mexico؛ فقط نفصل إذا جاء كفاصل مستقل
    if "|" in line:
        return [p.strip() for p in line.split("|") if p.strip()]
    if "*" in line:
        return [p.strip() for p in line.split("*") if p.strip()]
    if "×" in line:
        return [p.strip() for p in line.split("×") if p.strip()]
    if re.search(r"\s+[xX]\s+", line):
        return [p.strip() for p in re.split(r"\s+[xX]\s+", line) if p.strip()]
    return [line]


def _looks_like_match_line(line):
    p = _split_data_parts(line)
    return len(p) >= 2 and not re.search(r"\d+\s*[*×|\-–:]\s*\d+", line or "")


def _looks_like_result_line(line):
    if re.search(r"\d+\s*[*×|\-–:]\s*\d+", line or ""):
        return True
    p = _split_data_parts(line)
    if len(p) >= 4:
        return bool(re.search(r"^-?\d+$", p[1]) and re.search(r"^-?\d+$", p[2]))
    return False


def _maybe_day_and_rows(body_lines, kind="matches"):
    if not body_lines:
        return "اليوم", []
    first = body_lines[0]
    if kind == "results" and _looks_like_result_line(first):
        return "اليوم", body_lines
    if kind == "matches" and _looks_like_match_line(first):
        return "اليوم", body_lines
    return first, body_lines[1:]



EN_WEEKDAYS = {
    'السبت': 'SATURDAY', 'الاحد': 'SUNDAY', 'الأحد': 'SUNDAY', 'الاثنين': 'MONDAY', 'الإثنين': 'MONDAY',
    'الثلاثاء': 'TUESDAY', 'الاربعاء': 'WEDNESDAY', 'الأربعاء': 'WEDNESDAY', 'الخميس': 'THURSDAY',
    'الجمعة': 'FRIDAY', 'الاحد ': 'SUNDAY'
}
EN_MONTHS = {
    1: 'JANUARY', 2: 'FEBRUARY', 3: 'MARCH', 4: 'APRIL', 5: 'MAY', 6: 'JUNE',
    7: 'JULY', 8: 'AUGUST', 9: 'SEPTEMBER', 10: 'OCTOBER', 11: 'NOVEMBER', 12: 'DECEMBER'
}

def _format_design_date_text(raw):
    raw = normalize_name(str(raw or '').strip())
    if not raw:
        return ''
    m = re.match(r'^(\d{1,2})\/(\d{1,2})\/(\d{4})$', raw)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            dt = datetime(y, mo, d)
            return f"{dt.strftime('%A').upper()}, {dt.day} {EN_MONTHS.get(dt.month, dt.strftime('%B').upper())} {dt.year}"
        except Exception:
            return raw
    if raw in EN_WEEKDAYS:
        return EN_WEEKDAYS[raw]
    return raw.upper()


def _display_time_en(t):
    s = normalize_name(str(t or '').strip())
    upper = s.upper()
    if 'فجر' in s or 'ص' in s or 'AM' in upper:
        period = 'AM'
    elif 'م' in s or 'مساء' in s or 'PM' in upper:
        period = 'PM'
    else:
        period = ''
    tm = re.sub(r'(فجرًا|فجراً|فجرا|فجر|صباحًا|صباحا|ص|مساءً|مساء|م|AM|PM|am|pm)', '', s).strip()
    return f"{tm} {period}".strip()

# -------------------- Parsers V27 --------------------

def parse_matches_text(text):
    body = _body_lines_after_command(text)
    day_name, data_lines = _maybe_day_and_rows(body, "matches")
    matches = []
    for line in data_lines:
        parts = _split_data_parts(line)
        if len(parts) >= 3:
            matches.append((normalize_name(parts[0]), normalize_name(parts[1]), normalize_name(parts[2])))
        elif len(parts) == 2:
            # يدعم: فرنسا * البرتغال بدون وقت
            matches.append((normalize_name(parts[0]), normalize_name(parts[1]), ""))
        else:
            # يدعم: فرنسا - البرتغال - 8:00 م
            m = re.match(r"(.+?)\s*[-–]\s*(.+?)(?:\s*[-–]\s*(.+))?$", line)
            if m:
                matches.append((normalize_name(m.group(1)), normalize_name(m.group(2)), normalize_name(m.group(3) or "")))
    return normalize_name(day_name or "اليوم"), matches


def parse_match_results_design_text(text):
    body = _body_lines_after_command(text)
    day_name, data_lines = _maybe_day_and_rows(body, "results")
    results = []
    for line in data_lines:
        line = normalize_name(line)
        parts = _split_data_parts(line)
        if len(parts) >= 4:
            # يدعم: فريق * 2 * 1 * فريق أو فريق|2|1|فريق
            try:
                if re.fullmatch(r"\d+", parts[1]) and re.fullmatch(r"\d+", parts[2]):
                    results.append((normalize_name(parts[0]), int(parts[1]), int(parts[2]), normalize_name(parts[3])))
                    continue
            except Exception:
                pass
        # يدعم الصيغة المطلوبة: فرنسا 2 * 6 البرتغال
        m = re.match(r"(.+?)\s+(\d+)\s*[*×|\-–:]\s*(\d+)\s+(.+)$", line)
        if m:
            results.append((normalize_name(m.group(1)), int(m.group(2)), int(m.group(3)), normalize_name(m.group(4))))
            continue
        # يدعم: فرنسا 2 * البرتغال 6 بشكل احتياطي
        if len(parts) == 2:
            m1 = re.match(r"(.+?)\s+(\d+)\s*$", parts[0])
            m2 = re.match(r"^\s*(\d+)\s+(.+)$", parts[1])
            if m1 and m2:
                results.append((normalize_name(m1.group(1)), int(m1.group(2)), int(m2.group(1)), normalize_name(m2.group(2))))
    return normalize_name(day_name or "اليوم"), results


def parse_scorers_text(text):
    body = _body_lines_after_command(text)
    items = []
    for line in body:
        parts = _split_data_parts(line)
        name = team = ""
        goals = 0
        if len(parts) >= 3:
            # الجديد: ميسي * الأرجنتين * 3
            if re.fullmatch(r"\d+", parts[2]):
                name, team, goals = normalize_name(parts[0]), normalize_name(parts[1]), int(parts[2])
            # القديم: ميسي|3|الأرجنتين
            elif re.fullmatch(r"\d+", parts[1]):
                name, goals, team = normalize_name(parts[0]), int(parts[1]), normalize_name(parts[2])
        elif len(parts) == 2 and re.fullmatch(r"\d+", parts[1]):
            name, goals, team = normalize_name(parts[0]), int(parts[1]), ""
        if name and goals > 0:
            items.append((name, goals, team))
    return sorted(items, key=lambda x: (-x[1], x[0]))



def parse_multi_days_matches_text(text):
    body = _body_lines_after_command(text)
    sections = []
    current_date = None
    current_matches = []
    for line in body:
        line = normalize_name(line)
        if not line:
            continue
        if re.fullmatch(r'\d{1,2}/\d{1,2}/\d{4}', line):
            if current_date and current_matches:
                sections.append((current_date, current_matches))
            current_date = line
            current_matches = []
            continue
        parts = _split_data_parts(line)
        if len(parts) >= 3:
            current_matches.append((normalize_name(parts[0]), normalize_name(parts[1]), normalize_name(parts[2])))
        elif len(parts) == 2:
            current_matches.append((normalize_name(parts[0]), normalize_name(parts[1]), ''))
        else:
            m = re.match(r'(.+?)\s*[*|\-–]\s*(.+?)(?:\s*[*|\-–]\s*(.+))?$', line)
            if m:
                current_matches.append((normalize_name(m.group(1)), normalize_name(m.group(2)), normalize_name(m.group(3) or '')))
    if current_date and current_matches:
        sections.append((current_date, current_matches))
    return sections

def _extract_group_row(line):
    line = normalize_name(line)
    if not line:
        return None
    parts = _split_data_parts(line)
    if len(parts) >= 4:
        team = normalize_name(parts[0])
        nums = re.findall(r"[+-]?\d+", " ".join(parts[1:]))
        if len(nums) >= 3:
            return (team, int(nums[0]), int(nums[1]), int(nums[2]))
    # يدعم: اسكتلندا لعب 1 اهداف 1 نقاط 6
    m = re.match(r"(.+?)\s+لعب\s+([0-9]+)\s+(?:فارق|اهداف|أهداف|\+/-)\s+([+-]?[0-9]+)\s+نقاط\s+([0-9]+)\s*$", line)
    if m:
        return (normalize_name(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4)))
    # يدعم: اسكتلندا 1 +1 6
    m = re.match(r"(.+?)\s+([0-9]+)\s+([+-]?[0-9]+)\s+([0-9]+)\s*$", line)
    if m:
        return (normalize_name(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4)))
    return None


def parse_group_standing_text(text):
    body = _body_lines_after_command(text)
    if not body:
        return "المجموعة", []
    group_title = normalize_name(body[0])
    rows = []
    for line in body[1:]:
        row = _extract_group_row(line)
        if row:
            rows.append(row)
    rows.sort(key=lambda x: (x[3], x[2], x[0]), reverse=True)
    return group_title, rows


def parse_all_groups_text(text):
    body = _body_lines_after_command(text)
    groups = []
    current_title = None
    current_rows = []
    for line in body:
        row = _extract_group_row(line)
        if row and current_title:
            current_rows.append(row)
            continue
        # أي سطر ليس صف يعتبر عنوان مجموعة: A / المجموعة A
        if current_title and current_rows:
            current_rows.sort(key=lambda x: (x[3], x[2], x[0]), reverse=True)
            groups.append((current_title, current_rows))
            current_rows = []
        current_title = normalize_name(line)
    if current_title and current_rows:
        current_rows.sort(key=lambda x: (x[3], x[2], x[0]), reverse=True)
        groups.append((current_title, current_rows))
    return groups


# -------------------- Captions V27 --------------------

def build_design_matches_caption(day_name, matches):
    lines = ["🏆 مونديال المصيف 2026 🏆", f"🔥 مباريات اليوم ( {day_name} ) 🔥", ""]
    for a, b, t in matches:
        lines.append(f"{a} × {b}" + (f" — {t}" if t else ""))
    lines.append("")
    lines.append("المصيف ينقل لكم الحدث")
    return "\n".join(lines)


def build_design_results_caption(day_name, results):
    lines = ["🏆 نتائج مباريات اليوم 🏆", f"اليوم ( {day_name} )", ""]
    for a, sa, sb, b in results:
        lines.append(f"{a} {sa} - {sb} {b}")
    lines.append("")
    lines.append("المصيف ينقل لكم الحدث")
    return "\n".join(lines)


def build_match_results_caption(results):
    # للتوافق مع الأوامر القديمة التي لا تمرر اليوم
    lines = ["🏆 نتائج مباريات اليوم 🏆", ""]
    for a, sa, sb, b in results:
        lines.append(f"{a} {sa} - {sb} {b}")
    lines.append("\nالمصيف ينقل لكم الحدث")
    return "\n".join(lines)


def build_top_scorers_caption(items):
    lines = ["🏆 هدافين البطولة 🏆", ""]
    for i, (name, goals, team) in enumerate(sorted(items, key=lambda x: (-x[1], x[0])), start=1):
        team_part = f" — {team}" if team else ""
        lines.append(f"{i}. {name}{team_part} — {goals} أهداف")
    lines.append("\nالمصيف ينقل لكم الحدث")
    return "\n".join(lines)


# -------------------- New Look Backgrounds/Designs V27 --------------------

def _template_bg_path(style):
    name = "games_style2_bg.png" if int(style) == 2 else "games_style4_bg.png"
    return os.path.join("assets", "templates", name)


def _style4_clean_background(width=1200, height=1500):
    img = Image.new("RGB", (width, height), "#07153A")
    draw = ImageDraw.Draw(img)

    for y in range(height):
        t = y / max(1, height - 1)
        r = int(6 + 8 * t)
        g = int(18 + 18 * t)
        b = int(58 + 34 * t)
        draw.line((0, y, width, y), fill=(r, g, b))

    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)

    # وهج علوي ودائري يمين
    od.ellipse((110, -150, width - 130, 470), fill=(48, 95, 220, 70))
    od.ellipse((width * 0.62, 350, width * 1.12, 840), fill=(45, 96, 220, 55))
    od.ellipse((width * 0.70, 760, width * 1.18, 1280), fill=(32, 82, 205, 52))
    od.rectangle((width * 0.74, 450, width * 0.94, 690), fill=(60, 110, 230, 42))
    od.rectangle((width * 0.88, 550, width * 1.04, 890), fill=(48, 96, 210, 38))
    od.ellipse((width - 40, 420, width + 240, 720), fill=(35, 95, 220, 48))

    # لمسة خطوط/دوائر خفيفة يمين
    for off in [0, 115, 230, 345]:
        od.arc((830, 120 + off, 1160, 410 + off), 255, 105, fill=(89, 140, 255, 42), width=4)

    # ظل تمثال حرية مبسط يسار (بدون أي مربعات/خانات ثابتة)
    stat = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    sd = ImageDraw.Draw(stat)
    sx = int(width * 0.12)
    base_y = int(height * 0.78)
    blue = (20, 72, 180, 108)
    sd.polygon([
        (sx - 145, base_y + 210), (sx + 200, base_y + 210),
        (sx + 145, base_y - 210), (sx - 5, base_y - 255)
    ], fill=blue)
    sd.polygon([
        (sx - 70, base_y - 190), (sx + 60, base_y - 190),
        (sx + 35, base_y - 370), (sx - 45, base_y - 375)
    ], fill=(22, 82, 195, 98))
    sd.ellipse((sx - 45, base_y - 500, sx + 75, base_y - 375), fill=(22, 82, 195, 96))
    for ang in [-72, -45, -20, 10, 38, 62]:
        x2 = sx + 18 + int(120 * (ang / 72))
        sd.line((sx + 18, base_y - 470, x2, base_y - 620), fill=(22, 82, 195, 90), width=10)
    sd.line((sx - 105, base_y - 410, sx - 200, base_y - 705), fill=(22, 82, 195, 92), width=34)
    sd.ellipse((sx - 238, base_y - 782, sx - 155, base_y - 690), fill=(22, 82, 195, 95))
    sd.polygon([(sx - 198, base_y - 814), (sx - 235, base_y - 730), (sx - 160, base_y - 735)], fill=(22, 82, 195, 92))
    stat = stat.filter(ImageFilter.GaussianBlur(1))
    overlay = Image.alpha_composite(overlay, stat)

    overlay = overlay.filter(ImageFilter.GaussianBlur(6))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)

    # أقواس علوية خفيفة مثل المرجع
    for i in range(0, 55, 6):
        draw.arc((80 + i * 2, 34 + i, width - 85 - i * 2, 520 + i), 198, 336, fill="#5FA3FF44", width=2)

    return img, draw


def _draw_date_pill_only(draw, width, date_text):
    # شريط التاريخ الديناميكي فقط — بدون إعادة رسم عنوان الخلفية الأصلي.
    rounded_rect(draw, (width//2-240, 276, width//2+240, 334), radius=18, fill="#FBBF24", outline="#00000070", width=2)
    draw_text(draw, (width//2, 306), date_text, get_font(28), fill="#061633", max_width=450)


def _cover_style4_dynamic_areas(img, width, height, date_text):
    # نحافظ على خلفية التمثال والعناوين والرعاة وشعار السعودية،
    # ونغطي فقط تاريخ القالب القديم وخانات المباريات الثابتة.
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)

    # تغطية مستطيل التاريخ الأصفر القديم
    rounded_rect(od, (width//2-255, 266, width//2+255, 344), radius=22, fill="#08245FD8")

    # تغطية خانات المباريات الثابتة فقط مع ترك التمثال والعناصر الجانبية ظاهرة
    rounded_rect(od, (150, 350, 985, 1138), radius=38, fill="#061B58D8")

    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    draw = ImageDraw.Draw(img)
    _draw_date_pill_only(draw, width, date_text)
    return img, draw


def _newlook_canvas(style, title, sub_title, width=1200, height=1500):
    style = int(style)

    if style == 4:
        path = _template_bg_path(style)
        if os.path.exists(path):
            try:
                img = Image.open(path).convert("RGB").resize((width, height), Image.LANCZOS)
                return _cover_style4_dynamic_areas(img, width, height, sub_title)
            except Exception:
                pass
        # احتياط فقط إذا لم توجد الخلفية الأصلية
        img, draw = _style4_clean_background(width, height)
        _draw_games_header(draw, width, title, sub_title)
        return img, draw

    path = _template_bg_path(style)
    if os.path.exists(path):
        try:
            img = Image.open(path).convert("RGB").resize((width, height))
            draw = ImageDraw.Draw(img)
            overlay = Image.new("RGBA", (width, height), (0,0,0,0))
            od = ImageDraw.Draw(overlay)
            rounded_rect(od, (55, 45, width-55, 330), radius=30, fill="#06152F88")
            rounded_rect(od, (170, 350, width-170, height-265), radius=36, fill="#03133499")
            img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
            draw = ImageDraw.Draw(img)
        except Exception:
            img, draw = _games_day_background(width, height)
    else:
        img, draw = _games_day_background(width, height)
    _draw_games_header(draw, width, title, sub_title)
    return img, draw


def _newlook_card(draw, box, i=1, style=2):
    base_fill = "#0638A5" if int(style) == 2 else "#061E64"
    accent = v16_accent(i)
    rounded_rect(draw, box, radius=28, fill=base_fill, outline="#13B6EA" if int(style)==2 else "#3B82F6", width=3)
    x1,y1,x2,y2 = box
    draw.line((x1+10, y2-4, x2-10, y2-4), fill="#EF4444", width=4)
    draw.line((x1+90, y1+4, x2-90, y1+4), fill="#22C55E", width=3)
    draw.line((x1+310, y1+4, x1+375, y1+4), fill="#FBBF24", width=3)
    draw.arc((x1, y1, x1+64, y1+64), 180, 270, fill=accent, width=6)
    draw.arc((x2-64, y2-64, x2, y2), 0, 90, fill=accent, width=6)


def _newlook_layout(count):
    if count <= 1:
        return 225, 0, 505
    if count == 2:
        return 200, 34, 430
    if count == 3:
        return 178, 26, 405
    if count == 4:
        return 155, 21, 388
    return 128, 14, 370


def create_matches_newlook_image(day_name, matches, style=2):
    ensure_generated_dir()
    count = max(len(matches), 1)
    width, height = 1200, 1500
    img, draw = _newlook_canvas(style, "GAMES OF THE DAY", _format_design_date_text(day_name), width, height)
    row_h, gap, y = _newlook_layout(count)
    x1, x2 = 240, 960
    for i, (a, b, t) in enumerate(matches[:7], start=1):
        _newlook_card(draw, (x1, y, x2, y+row_h), i, style)
        cy = y + row_h//2
        flag_w = min(148, max(95, row_h - 34))
        # الفريق الأول يمين، الثاني يسار
        paste_flag(img, a, (x2-38-flag_w, cy-flag_w//2, x2-38, cy+flag_w//2))
        paste_flag(img, b, (x1+38, cy-flag_w//2, x1+38+flag_w, cy+flag_w//2))
        draw_text(draw, (x2-38-flag_w//2, y+row_h-25), _clean_display_name(a), get_font(25 if count>=4 else 29), fill="#FFFFFF", max_width=210)
        draw_text(draw, (x1+38+flag_w//2, y+row_h-25), _clean_display_name(b), get_font(25 if count>=4 else 29), fill="#FFFFFF", max_width=210)
        center = t or "VS"
        if t:
            if int(style) == 4:
                draw_text(draw, (width//2, cy), _display_time_en(t), get_font(40 if count>=4 else 48), fill="#FBBF24", max_width=230)
            else:
                tm, period = _ampm_from_time(t)
                draw_text(draw, (width//2, cy-12), tm, get_font(50 if count>=4 else 62), fill="#FBBF24")
                if period:
                    draw_text(draw, (width//2, cy+38), period, get_font(29 if count>=4 else 36), fill="#FBBF24")
        else:
            draw_text(draw, (width//2, cy), "VS", get_font(50), fill="#FBBF24")
        y += row_h + gap
    _draw_games_footer(draw, img, width, height)
    path = os.path.join(GENERATED_DIR, f"matches_style{style}_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path


def create_match_results_newlook_image(day_name, results, style=2):
    ensure_generated_dir()
    count = max(len(results), 1)
    width, height = 1200, 1500
    img, draw = _newlook_canvas(style, "MATCH RESULTS", _format_design_date_text(day_name), width, height)
    row_h, gap, y = _newlook_layout(count)
    x1, x2 = 240, 960
    for i, (a, sa, sb, b) in enumerate(results[:7], start=1):
        _newlook_card(draw, (x1, y, x2, y+row_h), i, style)
        cy = y + row_h//2
        flag_w = min(145, max(92, row_h - 38))
        paste_flag(img, a, (x2-38-flag_w, cy-flag_w//2, x2-38, cy+flag_w//2))
        paste_flag(img, b, (x1+38, cy-flag_w//2, x1+38+flag_w, cy+flag_w//2))
        draw_text(draw, (x2-38-flag_w//2, y+row_h-25), _clean_display_name(a), get_font(24 if count>=4 else 28), fill="#FFFFFF", max_width=210)
        draw_text(draw, (x1+38+flag_w//2, y+row_h-25), _clean_display_name(b), get_font(24 if count>=4 else 28), fill="#FFFFFF", max_width=210)
        rounded_rect(draw, (width//2-100, cy-44, width//2+100, cy+44), radius=20, fill="#FBBF24", outline="#00000088", width=2)
        draw_text(draw, (width//2, cy), f"{sa} - {sb}", get_font(52 if count>=4 else 60), fill="#061633")
        y += row_h + gap
    _draw_games_footer(draw, img, width, height)
    path = os.path.join(GENERATED_DIR, f"results_style{style}_{_safe_filename(day_name)}.png")
    img.save(path, quality=95)
    return path


def create_scorers_newlook_image(items, style=2):
    ensure_generated_dir()
    items = sorted(items, key=lambda x: (-x[1], x[0]))[:12]
    count = max(len(items), 1)
    width, height = 1200, 1500
    img, draw = _newlook_canvas(style, "TOP SCORERS", "هدافين البطولة", width, height)
    if count <= 3:
        row_h, gap, y = 122, 26, 430
    elif count <= 6:
        row_h, gap, y = 98, 18, 410
    else:
        row_h, gap, y = 74, 11, 395
    x1, x2 = 210, 990
    for i, (name, goals, team) in enumerate(items, start=1):
        _newlook_card(draw, (x1, y, x2, y+row_h), i, style)
        cy = y + row_h//2
        draw_text(draw, (x2-55, cy), str(i), get_font(32 if count>6 else 38), fill="#FBBF24")
        if team:
            fw = min(84, max(58, row_h-18))
            paste_flag(img, team, (x2-155, cy-fw//2, x2-155+fw, cy+fw//2))
        draw_text(draw, (width//2+55, cy), _clean_display_name(name), get_font(25 if count>6 else 32), fill="#FFFFFF", max_width=430)
        rounded_rect(draw, (x1+35, cy-34, x1+185, cy+34), radius=18, fill="#FBBF24", outline="#00000070", width=1)
        draw_text(draw, (x1+110, cy), str(goals), get_font(40 if count<=6 else 32), fill="#061633")
        draw_text(draw, (x1+250, cy), "أهداف", get_font(22 if count>6 else 26), fill="#FDE68A")
        y += row_h + gap
    _draw_games_footer(draw, img, width, height)
    path = os.path.join(GENERATED_DIR, f"scorers_style{style}.png")
    img.save(path, quality=95)
    return path


def create_group_newlook_image(group_title, rows, style=2):
    ensure_generated_dir()
    width, height = 1200, 1500
    title = clean_group_title_for_design(group_title)
    if not str(title).startswith("المجموعة"):
        title = f"المجموعة {title}"
    img, draw = _newlook_canvas(style, "GROUP STANDINGS", title, width, height)
    x1, x2 = 190, 1010
    y = 395
    rounded_rect(draw, (x1, y, x2, y+60), radius=20, fill="#061633DD", outline="#FFFFFF44", width=1)
    draw_text(draw, (x2-175, y+30), "المنتخب", get_font(24), fill="#FFFFFF")
    draw_text(draw, (x1+385, y+30), "لعب", get_font(23), fill="#FDE68A")
    draw_text(draw, (x1+245, y+30), "+/-", get_font(23), fill="#FDE68A")
    draw_text(draw, (x1+95, y+30), "نقاط", get_font(23), fill="#FDE68A")
    y += 78
    row_h, gap = 98, 14
    for i, (team, played, diff, pts) in enumerate(rows[:8], start=1):
        _newlook_card(draw, (x1, y, x2, y+row_h), i, style)
        cy = y + row_h//2
        draw_text(draw, (x2-48, cy), str(i), get_font(28), fill="#FBBF24")
        paste_flag(img, team, (x2-135, cy-32, x2-70, cy+32))
        draw_text(draw, (x2-285, cy), _clean_display_name(team), get_font(27), fill="#FFFFFF", max_width=250)
        draw_text(draw, (x1+385, cy), str(played), get_font(27), fill="#FFFFFF")
        draw_text(draw, (x1+245, cy), f"{int(diff):+d}", get_font(27), fill="#E5E7EB")
        draw_text(draw, (x1+95, cy), str(pts), get_font(34), fill="#FBBF24")
        y += row_h + gap
    _draw_games_footer(draw, img, width, height)
    path = os.path.join(GENERATED_DIR, f"group_style{style}_{_safe_filename(group_title)}.png")
    img.save(path, quality=95)
    return path


def create_all_groups_newlook_image(groups, style=2):
    ensure_generated_dir()
    width, height = 1800, 2400
    img, draw = _games_day_background(width, height)
    draw_text(draw, (width//2, 90), "MONDIAL AL MASEEF 2026", get_font(40), fill="#FFFFFF")
    draw_text(draw, (width//2, 170), "ALL GROUP STANDINGS", get_font(72), fill="#FFFFFF", max_width=width-160)
    draw_text(draw, (width//2, 235), "ترتيب جميع المجموعات", get_font(36), fill="#FBBF24")
    cols = 3
    margin_x, gap_x = 75, 35
    card_w = (width - 2*margin_x - (cols-1)*gap_x) // cols
    card_h = 470
    start_y = 320
    gap_y = 38
    for idx, (title, rows) in enumerate(groups[:12]):
        c = idx % cols
        r = idx // cols
        x = margin_x + c*(card_w+gap_x)
        y = start_y + r*(card_h+gap_y)
        rounded_rect(draw, (x, y, x+card_w, y+card_h), radius=28, fill="#0638A5EE", outline="#14B8F5", width=3)
        rounded_rect(draw, (x+18, y+18, x+card_w-18, y+68), radius=18, fill="#FBBF24", outline="#00000055", width=1)
        gt = clean_group_title_for_design(title)
        if not str(gt).startswith("المجموعة"):
            gt = f"المجموعة {gt}"
        draw_text(draw, (x+card_w//2, y+43), gt, get_font(26), fill="#061633", max_width=card_w-40)
        yy = y + 92
        for pos, (team, played, diff, pts) in enumerate(rows[:4], start=1):
            rounded_rect(draw, (x+20, yy, x+card_w-20, yy+72), radius=16, fill="#061633AA", outline="#FFFFFF22", width=1)
            cy = yy + 36
            draw_text(draw, (x+card_w-44, cy), str(pos), get_font(22), fill="#FBBF24")
            paste_flag(img, team, (x+card_w-125, cy-24, x+card_w-75, cy+24))
            draw_text(draw, (x+card_w-245, cy), _clean_display_name(team), get_font(21), fill="#FFFFFF", max_width=185)
            draw_text(draw, (x+155, cy), str(pts), get_font(26), fill="#FBBF24")
            draw_text(draw, (x+75, cy), f"{int(diff):+d}", get_font(21), fill="#E5E7EB")
            yy += 84
    draw_text(draw, (width//2, height-70), "المصيف ينقل لكم الحدث", get_font(36), fill="#FBBF24")
    path = os.path.join(GENERATED_DIR, f"all_groups_style{style}.png")
    img.save(path, quality=95)
    return path


# Override old style2 names to use V27 new look.
def create_matches_style2_image(day_name, matches):
    return create_matches_newlook_image(day_name, matches, 2)


def create_match_results_style2_image(day_name, results):
    return create_match_results_newlook_image(day_name, results, 2)


def create_scorers_style2_image(items):
    return create_scorers_newlook_image(items, 2)


def create_group_style2_image(group_title, rows):
    return create_group_newlook_image(group_title, rows, 2)


def create_all_groups_image(groups):
    return create_all_groups_newlook_image(groups, 2)


# -------------------- Style 3 frame designs for all design types --------------------

def create_scorers_frame_style_image(items):
    ensure_generated_dir()
    items = sorted(items, key=lambda x: (-x[1], x[0]))[:12]
    count = max(len(items), 1)
    width = 1200
    row_h = 92 if count <= 8 else 72
    gap = 14 if count <= 8 else 9
    height = max(760, 220 + count*(row_h+gap) + 120)
    img, draw = _frame_canvas(width, height)
    draw_text(draw, (width//2, 80), "مونديال المصيف 2026", get_font(50), fill="#FFFFFF")
    draw_text(draw, (width//2, 140), "هدافين البطولة", get_font(36), fill="#FDE68A")
    y = 215
    for i, (name, goals, team) in enumerate(items, start=1):
        x1, x2 = 110, width-110
        rounded_rect(draw, (x1, y, x2, y+row_h), radius=24, fill="#07110FCC", outline="#22C55E" if i==1 else "#2563EB", width=3)
        cy = y + row_h//2
        draw_text(draw, (x2-45, cy), str(i), get_font(28), fill="#FDE68A")
        if team:
            paste_flag(img, team, (x2-130, cy-30, x2-70, cy+30))
        draw_text(draw, (width//2+95, cy), _clean_display_name(name), get_font(28 if count<=8 else 23), fill="#FFFFFF", max_width=450)
        rounded_rect(draw, (x1+45, cy-30, x1+180, cy+30), radius=16, fill="#B8FFF0", outline="#FFFFFFAA", width=2)
        draw_text(draw, (x1+112, cy), str(goals), get_font(34 if count<=8 else 28), fill="#061633")
        y += row_h + gap
    draw_text(draw, (width//2, height-42), "المصيف ينقل لكم الحدث", get_font(28), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, "scorers_frame.png")
    img.save(path, quality=95)
    return path


def create_group_frame_style_image(group_title, rows):
    ensure_generated_dir()
    width = 1200
    count = max(len(rows), 1)
    row_h = 92 if count <= 8 else 74
    gap = 13 if count <= 8 else 8
    height = max(820, 245 + count*(row_h+gap) + 110)
    img, draw = _frame_canvas(width, height)
    title = clean_group_title_for_design(group_title)
    if not str(title).startswith("المجموعة"):
        title = f"المجموعة {title}"
    draw_text(draw, (width//2, 80), "مونديال المصيف 2026", get_font(50), fill="#FFFFFF")
    draw_text(draw, (width//2, 140), title, get_font(36), fill="#FDE68A")
    y = 215
    x1, x2 = 95, width-95
    rounded_rect(draw, (x1, y, x2, y+56), radius=20, fill="#061633", outline="#22C55E", width=2)
    draw_text(draw, (940, y+28), "المنتخب", get_font(24), fill="#FFFFFF")
    draw_text(draw, (530, y+28), "لعب", get_font(22), fill="#FDE68A")
    draw_text(draw, (385, y+28), "+/-", get_font(22), fill="#FDE68A")
    draw_text(draw, (220, y+28), "نقاط", get_font(22), fill="#FDE68A")
    y += 76
    for i, (team, played, diff, pts) in enumerate(rows[:10], start=1):
        rounded_rect(draw, (x1, y, x2, y+row_h), radius=24, fill="#07110FCC", outline="#22C55E" if i==1 else "#2563EB", width=2)
        cy = y + row_h//2
        draw_text(draw, (1045, cy), str(i), get_font(25), fill="#FDE68A")
        paste_flag(img, team, (940, cy-30, 1000, cy+30))
        draw_text(draw, (760, cy), _clean_display_name(team), get_font(27 if count<=8 else 23), fill="#FFFFFF", max_width=310)
        draw_text(draw, (530, cy), str(played), get_font(25), fill="#FFFFFF")
        draw_text(draw, (385, cy), f"{int(diff):+d}", get_font(25), fill="#FFFFFF")
        draw_text(draw, (220, cy), str(pts), get_font(32), fill="#FDE68A")
        y += row_h + gap
    draw_text(draw, (width//2, height-42), "المصيف ينقل لكم الحدث", get_font(28), fill="#FFFFFF")
    path = os.path.join(GENERATED_DIR, f"group_frame_{_safe_filename(group_title)}.png")
    img.save(path, quality=95)
    return path


def create_all_groups_frame_style_image(groups):
    # ستايل 3 لجميع المجموعات بنفس روح الإطار لكن أوسع.
    return create_all_groups_newlook_image(groups, 3)



def create_multi_days_matches_image(schedule_blocks, style=4):
    ensure_generated_dir()
    width, height = 1800, 2400
    img, draw = _games_day_background(width, height)
    draw_text(draw, (width//2, 90), "MONDIAL AL MASEEF 2026", get_font(40), fill="#FFFFFF")
    draw_text(draw, (width//2, 170), "MATCH SCHEDULE", get_font(72), fill="#FFFFFF", max_width=width-160)
    draw_text(draw, (width//2, 235), "جدول المباريات", get_font(36), fill="#FBBF24")
    cols = 2
    margin_x, gap_x = 90, 45
    card_w = (width - 2*margin_x - gap_x) // cols
    card_h = 650
    start_y = 320
    gap_y = 42
    for idx, (date_txt, matches) in enumerate(schedule_blocks[:6]):
        c = idx % cols
        r = idx // cols
        x = margin_x + c*(card_w+gap_x)
        y = start_y + r*(card_h+gap_y)
        rounded_rect(draw, (x, y, x+card_w, y+card_h), radius=28, fill="#0638A5EE", outline="#14B8F5", width=3)
        rounded_rect(draw, (x+18, y+18, x+card_w-18, y+82), radius=18, fill="#FBBF24", outline="#00000055", width=1)
        draw_text(draw, (x+card_w//2, y+50), _format_design_date_text(date_txt), get_font(26), fill="#061633", max_width=card_w-40)
        yy = y + 108
        for a, b, t in matches[:6]:
            rounded_rect(draw, (x+20, yy, x+card_w-20, yy+76), radius=16, fill="#061633AA", outline="#FFFFFF22", width=1)
            cy = yy + 38
            paste_flag(img, a, (x+card_w-120, cy-25, x+card_w-68, cy+25))
            paste_flag(img, b, (x+68, cy-25, x+120, cy+25))
            draw_text(draw, (x+card_w-225, cy), _clean_display_name(a), get_font(20), fill="#FFFFFF", max_width=165)
            draw_text(draw, (x+225, cy), _clean_display_name(b), get_font(20), fill="#FFFFFF", max_width=165)
            if t:
                tm, period = _ampm_from_time(t)
                mid = tm if not period else f"{tm} {period}"
            else:
                mid = 'VS'
            draw_text(draw, (x+card_w//2, cy), mid, get_font(20), fill="#FBBF24", max_width=160)
            yy += 88
    draw_text(draw, (width//2, height-70), "المصيف ينقل لكم الحدث", get_font(36), fill="#FBBF24")
    path = os.path.join(GENERATED_DIR, f"multi_days_style{style}.png")
    img.save(path, quality=95)
    return path

async def multi_days_matches_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        style = command_style(update.message.text, 4)
        blocks = parse_multi_days_matches_text(update.message.text)
        if not blocks:
            await update.message.reply_text("اكتبها كذا:\n/مباريات_الأيام 4\n24/06/2026\nالهلال * النصر * 8:00 م\nالهلال * الشباب * 10:00 م\n\n25/06/2026\nالأهلي * الاتحاد * 9:00 م")
            return
        path = create_multi_days_matches_image(blocks, style)
        await send_photo_path(update, path, "جدول مباريات عدة أيام ✅")
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم مباريات الأيام ❌\n{e}")

# -------------------- Render by style --------------------

def render_matches_by_style(day_name, matches, style, auto=False):
    style = int(style)
    if style == 1:
        return create_matches_template_image(day_name, matches, use_template=not auto)
    if style == 2:
        return create_matches_newlook_image(day_name, matches, 2)
    if style == 3:
        return create_match_frame_style_image(day_name, matches, False)
    return create_matches_newlook_image(day_name, matches, 4)


def render_results_by_style(day_name, results, style, auto=False):
    style = int(style)
    if style == 1:
        return create_match_results_template_image(day_name, results, use_template=not auto)
    if style == 2:
        return create_match_results_newlook_image(day_name, results, 2)
    if style == 3:
        return create_match_frame_style_image(day_name, results, True)
    return create_match_results_newlook_image(day_name, results, 4)


def render_scorers_by_style(items, style, auto=False):
    style = int(style)
    if style == 1:
        return create_top_scorers_template_image(items, use_template=not auto)
    if style == 2:
        return create_scorers_newlook_image(items, 2)
    if style == 3:
        return create_scorers_frame_style_image(items)
    return create_scorers_newlook_image(items, 4)


def render_group_by_style(group_title, rows, style, auto=False):
    style = int(style)
    if style == 1:
        return create_group_standing_image(group_title, rows, use_template=not auto)
    if style == 2:
        return create_group_newlook_image(group_title, rows, 2)
    if style == 3:
        return create_group_frame_style_image(group_title, rows)
    return create_group_newlook_image(group_title, rows, 4)


def render_all_groups_by_style(groups, style, auto=False):
    style = int(style)
    if style == 1:
        return create_all_groups_newlook_image(groups, 1)
    if style == 3:
        return create_all_groups_frame_style_image(groups)
    if style == 4:
        return create_all_groups_newlook_image(groups, 4)
    return create_all_groups_newlook_image(groups, 2)


# -------------------- Commands V27 --------------------

async def set_style_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    nums = get_numbers(update.message.text)
    if not nums or nums[0] not in (1, 2, 3, 4):
        await update.message.reply_text("اكتبها كذا:\n/اعتماد_ستايل 4\n\nالخيارات: 1، 2، 3، 4")
        return
    style = save_design_style(nums[0])
    await update.message.reply_text(f"تم اعتماد الستايل {style} ✅\n{STYLE_NAMES[style]}")


async def get_style_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    style = current_design_style()
    await update.message.reply_text(f"الستايل الحالي: {style} ✅\n{STYLE_NAMES.get(style, '')}")


async def matches_command_v27(update: Update, context: ContextTypes.DEFAULT_TYPE, forced_style=None, auto=False):
    try:
        style = command_style(update.message.text, forced_style)
        day_name, matches = parse_matches_text(update.message.text)
        if not matches:
            await update.message.reply_text("اكتبها كذا:\n/مباريات 4\nالسابع\nفرنسا * البرتغال * 8:00 م\nإنجلترا * كرواتيا * 11:00 م")
            return
        path = render_matches_by_style(day_name, matches, style, auto=auto)
        await send_photo_path(update, path, build_design_matches_caption(day_name, matches))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم المباريات ❌\n{e}")


async def results_command_v27(update: Update, context: ContextTypes.DEFAULT_TYPE, forced_style=None, auto=False):
    try:
        style = command_style(update.message.text, forced_style)
        day_name, results = parse_match_results_design_text(update.message.text)
        if not results:
            await update.message.reply_text("اكتبها كذا:\n/انتهت 4\nالسابع\nفرنسا 2 * 6 البرتغال\nإنجلترا 3 * 0 كرواتيا")
            return
        path = render_results_by_style(day_name, results, style, auto=auto)
        await send_photo_path(update, path, build_design_results_caption(day_name, results))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم نتائج المباريات ❌\n{e}")


async def scorers_command_v27(update: Update, context: ContextTypes.DEFAULT_TYPE, forced_style=None, auto=False):
    try:
        style = command_style(update.message.text, forced_style)
        items = parse_scorers_text(update.message.text)
        if not items:
            await update.message.reply_text("اكتبها كذا:\n/هدافين 4\nميسي * الأرجنتين * 3\nرونالدو * البرتغال * 6")
            return
        path = render_scorers_by_style(items, style, auto=auto)
        await send_photo_path(update, path, build_top_scorers_caption(items))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم الهدافين ❌\n{e}")


async def group_command_v27(update: Update, context: ContextTypes.DEFAULT_TYPE, forced_style=None, auto=False):
    try:
        style = command_style(update.message.text, forced_style)
        group_title, rows = parse_group_standing_text(update.message.text)
        if not rows:
            await update.message.reply_text("اكتبها كذا:\n/مجموعة 4\nC\nاسكتلندا * لعب 1 * فارق +1 * نقاط 6\nالمغرب * لعب 1 * فارق 0 * نقاط 1")
            return
        path = render_group_by_style(group_title, rows, style, auto=auto)
        await send_photo_path(update, path, build_group_standing_caption(group_title, rows))
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم ترتيب المجموعة ❌\n{e}")


async def all_groups_command_v27(update: Update, context: ContextTypes.DEFAULT_TYPE, forced_style=None, auto=False):
    try:
        style = command_style(update.message.text, forced_style)
        groups = parse_all_groups_text(update.message.text)
        if not groups:
            await update.message.reply_text("اكتبها كذا:\n/كل_المجموعات 4\nA\nالمكسيك * 1 * +2 * 3\nكوريا الجنوبية * 1 * +1 * 3\n\nB\nكندا * 1 * 0 * 1")
            return
        path = render_all_groups_by_style(groups, style, auto=auto)
        await send_photo_path(update, path, "ترتيب جميع المجموعات ✅")
    except Exception as e:
        await update.message.reply_text(f"تعذر تصميم جميع المجموعات ❌\n{e}")


# Wrappers عشان MessageHandler يستدعي دالة بدون بارامترات إضافية
async def short_matches_command(update, context):
    return await matches_command_v27(update, context)
async def short_matches_auto_command(update, context):
    return await matches_command_v27(update, context, auto=True)
async def short_results_command(update, context):
    return await results_command_v27(update, context)
async def short_results_auto_command(update, context):
    return await results_command_v27(update, context, auto=True)
async def short_scorers_command(update, context):
    return await scorers_command_v27(update, context)
async def short_scorers_auto_command(update, context):
    return await scorers_command_v27(update, context, auto=True)
async def short_group_command(update, context):
    return await group_command_v27(update, context)
async def short_group_auto_command(update, context):
    return await group_command_v27(update, context, auto=True)
async def short_all_groups_command(update, context):
    return await all_groups_command_v27(update, context)
async def short_all_groups_auto_command(update, context):
    return await all_groups_command_v27(update, context, auto=True)

# Old commands routed to the new 1-4 style system.
async def design_matches_template_command(update, context):
    return await matches_command_v27(update, context, auto=False)
async def design_matches_auto_command(update, context):
    return await matches_command_v27(update, context, auto=True)
async def design_match_results_template_command(update, context):
    return await results_command_v27(update, context, auto=False)
async def design_match_results_auto_command(update, context):
    return await results_command_v27(update, context, auto=True)
async def design_group_standing_template_command(update, context):
    return await group_command_v27(update, context, auto=False)
async def design_group_standing_auto_command(update, context):
    return await group_command_v27(update, context, auto=True)
async def design_scorers_template_command(update, context):
    return await scorers_command_v27(update, context, auto=False)
async def design_scorers_auto_command(update, context):
    return await scorers_command_v27(update, context, auto=True)

async def design_matches_style2_command(update, context):
    return await matches_command_v27(update, context, forced_style=2)
async def design_match_results_style2_command(update, context):
    return await results_command_v27(update, context, forced_style=2)
async def design_scorers_style2_command(update, context):
    return await scorers_command_v27(update, context, forced_style=2)
async def design_group_style2_command(update, context):
    return await group_command_v27(update, context, forced_style=2)
async def design_matches_frame_command(update, context):
    return await matches_command_v27(update, context, forced_style=3)
async def design_results_frame_command(update, context):
    return await results_command_v27(update, context, forced_style=3)
async def design_all_groups_command(update, context):
    return await all_groups_command_v27(update, context)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "البوت جاهز ✅\n\n"
        "الأوامر الأساسية:\n"
        "/اضافه 5\n/نتائج 5\n/اعتماد_نتائج 5\n/احصائيات\n/احصائيات 1 6\n/الترتيب_العام\n/الترتيب_العام 1 6\n/ترتيب_نص\n\n"
        "أوامر الصور والتقارير:\n"
        "/صورة_اليوم 6\n/صورة_الترتيب 1 6\n/صورة_الاساطير 1 6\n/صورة_احصائيات 1 6 لوحة عامة\n/صور_الاحصائيات 1 6\n"
        "/بطاقة فارس سالم\n/تقرير_الفترة 1 4\n/تفعيل_الصور_التلقائية\n/إيقاف_الصور_التلقائية\n\n"
        "أوامر الستايل:\n"
        "/اعتماد_ستايل 1\n/اعتماد_ستايل 2\n/اعتماد_ستايل 3\n/اعتماد_ستايل 4\n/الستايل\n\n"
        "أوامر التصاميم المختصرة:\n"
        "/مباريات 4\n/انتهت 4\n/مجموعة 4\n/هدافين 4\n/كل_المجموعات 4\n\n"
        "الأوامر التلقائية المختصرة:\n"
        "/مباريات_تلقائي 4\n/انتهت_تلقائي 4\n/مجموعة_تلقائي 4\n/هدافين_تلقائي 4\n/كل_المجموعات_تلقائي 4\n\n"
        "أوامر التصاميم القديمة باقية وتدعم 1-4:\n"
        "/تصميم_مباريات 4\n/تصميم_مباريات_تلقائي 4\n/تصميم_نتائج_مباريات 4\n/تصميم_نتائج_مباريات_تلقائي 4\n"
        "/تصميم_ترتيب_مجموعة 4\n/تصميم_ترتيب_مجموعة_تلقائي 4\n/تصميم_هدافين 4\n/تصميم_هدافين_تلقائي 4\n"
        "/تصميم_مباريات_ستايل2\n/تصميم_نتائج_مباريات_ستايل2\n/تصميم_هدافين_ستايل2\n/تصميم_ترتيب_مجموعة_ستايل2\n"
        "/تصميم_مباريات_اطار\n/تصميم_نتائج_مباريات_اطار\n/تصميم_جميع_المجموعات 4\n\n"
        "أوامر الكأس:\n"
        "/بدء_الكاس 7\n/حالة_الكاس\n/مواجهات_الكاس\n/نتائج_الكاس\n/إعادة_الكاس_من 7\n/الغاء_الكاس تأكيد\n\n"
        "أوامر الفحص والنشر:\n"
        "/الأيام\n/فحص 5\n/مشاركين 5\n/اسطورة 5\n/مقارنة 4 5\n/اعلان_اليوم 5\n/ملخص_اليوم 5\n\n"
        "أوامر الاستيراد والنسخ:\n"
        "/استيراد_ملف\n/اعتماد_استيراد\n/إلغاء_استيراد\n/نسخة_احتياطية\n/استرجاع_نسخة\n/تنظيف_الأيام\n/تنظيف_الملفات\n\n"
        "أوامر الأمان:\n"
        "/مسح_نتائج 5\n/مسح_يوم 5\n/مسح_الكل تأكيد\n/استرجاع_آخر\n/قفل_يوم 5\n/فتح_يوم 5\n/معرفي\n\n"
        "مهم: /نتائج للفانتزي فقط، و/انتهت لتصميم نتائج المباريات."
    )


# Main V27 — يعيد تسجيل كل الأوامر بالترتيب الصحيح.
def main():
    if not TOKEN:
        raise RuntimeError("ضع توكن البوت في متغير البيئة BOT_TOKEN")
    ensure_flags_assets()
    ensure_design_assets()
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/start"), start))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/(?:من_انا|معرفي)"), who_am_i))
    app.add_handler(MessageHandler(filters.Document.ALL, remember_last_file))

    # صور وتقارير
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تفعيل_الصور_التلقائية"), admin_only(enable_auto_images)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إيقاف_الصور_التلقائية"), admin_only(disable_auto_images)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/ايقاف_الصور_التلقائية"), admin_only(disable_auto_images)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صورة_اليوم"), daily_image_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صورة_الترتيب"), overall_image_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صورة_الاساطير"), legends_image_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صورة_احصائيات"), dashboard_sheet_image_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/صور_الاحصائيات"), all_dashboard_images_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/بطاقة"), participant_card_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تقرير_الفترة"), period_report_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعلان_اليوم"), announcement_day_command))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/ملخص_اليوم"), summary_day_command))

    # ستايل وتصاميم مختصرة — الأطول قبل الأقصر
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعتماد_ستايل(?:\s|$)"), admin_only(set_style_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الستايل(?:\s|$)"), admin_only(get_style_command)))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مباريات_تلقائي(?:\s|$)"), admin_only(short_matches_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/انتهت_تلقائي(?:\s|$)"), admin_only(short_results_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مجموعة_تلقائي(?:\s|$)"), admin_only(short_group_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/هدافين_تلقائي(?:\s|$)"), admin_only(short_scorers_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/كل_المجموعات_تلقائي(?:\s|$)"), admin_only(short_all_groups_auto_command)))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مباريات(?:\s|$)"), admin_only(short_matches_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/انتهت(?:\s|$)"), admin_only(short_results_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/كل_المجموعات(?:\s|$)"), admin_only(short_all_groups_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مجموعة(?:\s|$)"), admin_only(short_group_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/هدافين(?:\s|$)"), admin_only(short_scorers_command)))

    # أوامر التصميم القديمة + ستايلات قديمة
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_مباريات_ستايل2(?:\s|$)"), admin_only(design_matches_style2_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_نتائج_مباريات_ستايل2(?:\s|$)"), admin_only(design_match_results_style2_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_هدافين_ستايل2(?:\s|$)"), admin_only(design_scorers_style2_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_ترتيب_مجموعة_ستايل2(?:\s|$)"), admin_only(design_group_style2_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_مباريات_اطار(?:\s|$)"), admin_only(design_matches_frame_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_نتائج_مباريات_اطار(?:\s|$)"), admin_only(design_results_frame_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_جميع_المجموعات(?:\s|$)"), admin_only(design_all_groups_command)))

    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_مباريات_تلقائي(?:\s|$)"), admin_only(design_matches_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_نتائج_مباريات_تلقائي(?:\s|$)"), admin_only(design_match_results_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_ترتيب_مجموعة_تلقائي(?:\s|$)"), admin_only(design_group_standing_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_هدافين_تلقائي(?:\s|$)"), admin_only(design_scorers_auto_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_مباريات(?:\s|$)"), admin_only(design_matches_template_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_نتائج_مباريات(?:\s|$)"), admin_only(design_match_results_template_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_ترتيب_مجموعة(?:\s|$)"), admin_only(design_group_standing_template_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تصميم_هدافين(?:\s|$)"), admin_only(design_scorers_template_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج_مباريات_اليوم(?:\s|$)"), admin_only(design_match_results_template_command)))

    # استيراد ونسخ
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استيراد_ملف"), admin_only(import_excel_file)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعتماد_استيراد"), admin_only(approve_import)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إلغاء_استيراد"), admin_only(cancel_import)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الغاء_استيراد"), admin_only(cancel_import)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نسخة_احتياطية"), admin_only(backup_zip)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استرجاع_نسخة"), admin_only(restore_backup_zip)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تنظيف_الأيام"), admin_only(clean_days)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تنظيف_الايام"), admin_only(clean_days)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/تنظيف_الملفات"), admin_only(clean_temp_files)))

    # الكأس
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/بدء_الكاس(?:\s|$)"), admin_only(start_cup_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/حالة_الكاس(?:\s|$)"), admin_only(cup_status_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج_الكاس(?:\s|$)"), admin_only(cup_results_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مواجهات_الكاس(?:\s|$)"), admin_only(cup_matches_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إعادة_الكاس_من(?:\s|$)"), admin_only(reset_cup_from_day_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعادة_الكاس_من(?:\s|$)"), admin_only(reset_cup_from_day_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الغاء_الكاس(?:\s|$)"), admin_only(cancel_cup_command)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/إلغاء_الكاس(?:\s|$)"), admin_only(cancel_cup_command)))

    # فانتزي أساسي
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اضافه"), admin_only(add_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اعتماد_نتائج(?:\s|$)"), admin_only(approve_results_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/نتائج"), admin_only(results_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/الترتيب_العام"), overall))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/ترتيب_نص"), ranking_text))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/?(?:احصائيات|إحصائيات)(?:\s|$)"), dashboard))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/(الأيام|الايام)"), list_days))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/فحص"), inspect_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مشاركين"), participants_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/اسطورة"), legend_day))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مقارنة"), compare_days))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_الكل"), admin_only(clear_all)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_يوم"), admin_only(clear_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/مسح_نتائج"), admin_only(clear_results)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/استرجاع_آخر"), admin_only(restore_last)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/قفل_يوم"), admin_only(lock_day)))
    app.add_handler(MessageHandler(filters.TEXT & filters.Regex(r"^/فتح_يوم"), admin_only(unlock_day)))
    app.run_polling()


if __name__ == "__main__":
    main()