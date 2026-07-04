# -*- coding: utf-8 -*-
"""
بوت مسابقة المصيف اليدوي — نسخة من الصفر
- لا يسحب بيانات من أي موقع.
- لا يوجد بث مباشر/ESPN/نتائج/إشعارات.
- يعتمد فقط على Telegram + ملف contest_data.json المحلي.
"""

import os
import json
import hashlib
import logging
import re
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

try:
    from PIL import Image, ImageDraw, ImageFont
except Exception as e:
    Image = None
    ImageDraw = None
    ImageFont = None

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
except Exception:
    arabic_reshaper = None
    get_display = None

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("almaseef-contests")

TOKEN = (
    os.getenv("BOT_TOKEN")
    or os.getenv("TELEGRAM_BOT_TOKEN")
    or os.getenv("TELEGRAM_TOKEN")
    or os.getenv("TOKEN")
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = Path(os.getenv("CONTEST_DATA_FILE", BASE_DIR / "contest_data.json"))
GENERATED_DIR = Path(os.getenv("GENERATED_DIR", BASE_DIR / "generated"))
GENERATED_DIR.mkdir(parents=True, exist_ok=True)

PASSWORD = "الملحق"

DEFAULT_DATA = {
    "schema_version": "almaseef_manual_v1",
    "departed_teams": [],
    "contests": {
        "abokhaled": [
            {"name":"أبو خالد","team":"إنجلترا"},
            {"name":"سلمان أحمد","team":"إسبانيا"},
            {"name":"أبوداحم","team":"البرازيل"},
            {"name":"نواف فارس","team":"فرنسا"},
            {"name":"أبو نايف","team":"البرتغال"},
            {"name":"أبو راكان","team":"هولندا"},
            {"name":"مشعل","team":"البرتغال"},
            {"name":"نايف","team":"هولندا"},
            {"name":"محمد عبدالرحمن","team":"البرتغال"},
            {"name":"سلطان","team":"فرنسا"},
            {"name":"خالد","team":"إسبانيا"},
            {"name":"عادل","team":"البرتغال"},
            {"name":"أبو عبدالله","team":"السعودية"},
            {"name":"فهد فارس","team":"ألمانيا"},
            {"name":"زياد","team":"السعودية"},
            {"name":"فارس سالم","team":"فرنسا"},
            {"name":"مشاري عبدالعزيز","team":"إسبانيا"},
            {"name":"طلال عبدالله","team":"الأرجنتين"},
            {"name":"عبدالعزيز","team":"البرتغال"},
            {"name":"أبوتركي","team":"نيوزيلندا"},
            {"name":"أبوفارس","team":"إسبانيا"},
            {"name":"عبدالله","team":"فرنسا"},
            {"name":"أبو طلال","team":"ساحل العاج"},
            {"name":"خالد","team":"إسبانيا"},
            {"name":"يزيد","team":"إنجلترا"},
            {"name":"أبو يزيد","team":"إنجلترا"},
        ],
        "aboyaser": [
            {"name":"أبو فارس","team":"البرازيل","player":"ديمبلي"},
            {"name":"أبو هوى","team":"فرنسا","player":"يامال"},
            {"name":"تركي محسن","team":"البرازيل","player":"يامال"},
            {"name":"زيكا","team":"البرتغال","player":"مبابي"},
            {"name":"طلال عبدالله","team":"إنجلترا","player":"يامال"},
            {"name":"عبدالله إبراهيم","team":"فرنسا","player":"مبابي"},
            {"name":"بدران","team":"فرنسا","player":"أوليسي"},
            {"name":"فهد فارس","team":"ألمانيا","player":"يامال"},
            {"name":"أبو صنت","team":"السعودية","player":"سعود عبدالحميد"},
            {"name":"يزيد","team":"إسبانيا","player":"مبابي"},
            {"name":"سلطان أحمد","team":"فرنسا","player":"مبابي"},
            {"name":"نواف فارس","team":"إنجلترا","player":"هاري كين"},
            {"name":"الأمير","team":"إسبانيا","player":"يامال"},
            {"name":"فارس سالم","team":"فرنسا","player":"يامال"},
            {"name":"مشاري عبدالعزيز","team":"إسبانيا","player":"يامال"},
            {"name":"هندسة","team":"البرازيل","player":"يامال"},
            {"name":"سلمان أحمد","team":"إسبانيا","player":"يامال"},
            {"name":"نايف حمود","team":"الأرجنتين","player":"أبو محمد"},
            {"name":"عبدالرحمن سالم","team":"البرازيل","player":"مبابي"},
            {"name":"محمد محسن","team":"البرازيل","player":"بيدرو"},
            {"name":"أبو شنب","team":"فرنسا","player":"مبابي"},
            {"name":"ممدوح غزاي","team":"إسبانيا","player":"فيرتز"},
            {"name":"جلعده","team":"إسبانيا","player":"مبابي"},
            {"name":"خالد عبدالرحمن","team":"إسبانيا","player":"يامال"},
            {"name":"محمد عبدالرحمن","team":"البرتغال","player":"مبابي"},
            {"name":"سلطان رباح","team":"إسبانيا","player":"يامال"},
            {"name":"عادل","team":"البرتغال","player":"كين"},
        ],
    },
}

TEAM_ALIASES = {
    "اسبانيا":"إسبانيا", "أسبانيا":"إسبانيا", "إسبانيا":"إسبانيا",
    "السعوديه":"السعودية", "السعودية":"السعودية",
    "انجلترا":"إنجلترا", "إنجلترا":"إنجلترا", "انقلترا":"إنجلترا",
    "الارجنتين":"الأرجنتين", "ارجنتين":"الأرجنتين", "الأرجنتين":"الأرجنتين",
    "المانيا":"ألمانيا", "ألمانيا":"ألمانيا",
    "البرازيل":"البرازيل", "برازيل":"البرازيل", "برازيـل":"البرازيل",
    "البرتغال":"البرتغال", "برتغال":"البرتغال",
    "نيوزيلندا":"نيوزيلندا", "نيوزلندا":"نيوزيلندا",
    "ساحل العاج":"ساحل العاج", "كوت ديفوار":"ساحل العاج",
    "هولندا":"هولندا", "هولنده":"هولندا",
    "فرنسا":"فرنسا", "المغرب":"المغرب",
}

CONTEST_NAMES = {
    "abokhaled": "مسابقة أبوخالد",
    "aboyaser": "مسابقة أبوياسر",
}

# ---------- البيانات ----------

def _deepcopy_default():
    return json.loads(json.dumps(DEFAULT_DATA, ensure_ascii=False))


def load_data() -> dict:
    if DATA_FILE.exists():
        try:
            data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("data is not dict")
            # أي ملف قديم من نسخ سابقة لا نستخدمه حتى لا يخلط البيانات.
            if data.get("schema_version") != "almaseef_manual_v1":
                backup = DATA_FILE.with_suffix(f".old.{int(datetime.now().timestamp())}.json")
                try:
                    DATA_FILE.rename(backup)
                except Exception:
                    pass
                data = _deepcopy_default()
                save_data(data)
                return data
            data.setdefault("departed_teams", [])
            data.setdefault("contests", {})
            data["contests"].setdefault("abokhaled", [])
            data["contests"].setdefault("aboyaser", [])
            return data
        except Exception as e:
            backup = DATA_FILE.with_suffix(f".broken.{int(datetime.now().timestamp())}.json")
            try:
                DATA_FILE.rename(backup)
            except Exception:
                pass
            log.warning("تعذر قراءة البيانات، تم إنشاء ملف جديد: %s", e)
    data = _deepcopy_default()
    save_data(data)
    return data


def save_data(data: dict) -> None:
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp = DATA_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(DATA_FILE)


def data_hash(data: dict) -> str:
    raw = json.dumps(data, ensure_ascii=False, sort_keys=True)
    return hashlib.md5(raw.encode("utf-8")).hexdigest()[:10]

# ---------- النصوص والتطبيع ----------

def normalize_ar(text: str) -> str:
    s = str(text or "").strip()
    s = s.replace("ـ", "")
    s = re.sub(r"[\u064B-\u065F\u0670]", "", s)
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ى", "ي").replace("ة", "ه")
    s = re.sub(r"\s+", " ", s)
    return s.lower().strip()


def canonical_team(team: str) -> str:
    s = str(team or "").strip()
    if not s:
        return ""
    direct = TEAM_ALIASES.get(s)
    if direct:
        return direct
    n = normalize_ar(s)
    for k, v in TEAM_ALIASES.items():
        if normalize_ar(k) == n:
            return v
    return s


def team_key(team: str) -> str:
    return normalize_ar(canonical_team(team))


def ar(text: str) -> str:
    text = str(text or "")
    if arabic_reshaper and get_display:
        try:
            return get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
    return text


def now_riyadh() -> str:
    try:
        dt = datetime.now(ZoneInfo("Asia/Riyadh"))
    except Exception:
        dt = datetime.now()
    return dt.strftime("%Y-%m-%d %H:%M")


def contest_rows(data: dict, kind: str) -> list:
    return data.get("contests", {}).get(kind, []) or []


def departed_keys(data: dict) -> set:
    return {str(x) for x in data.get("departed_teams", []) if x}


def is_departed(data: dict, team: str) -> bool:
    return team_key(team) in departed_keys(data)


def all_selected_teams(data: dict, only_departed: bool = False) -> list[tuple[str, str]]:
    seen = {}
    for kind in ("abokhaled", "aboyaser"):
        for row in contest_rows(data, kind):
            t = canonical_team(row.get("team", ""))
            k = team_key(t)
            if not k:
                continue
            seen.setdefault(k, t)
    if only_departed:
        dep = departed_keys(data)
        seen = {k: v for k, v in seen.items() if k in dep}
    return sorted(seen.items(), key=lambda kv: kv[1])


def participants_by_team(data: dict, team_k: str) -> dict:
    res = {"abokhaled": [], "aboyaser": []}
    for kind in res:
        for row in contest_rows(data, kind):
            if team_key(row.get("team", "")) == team_k:
                res[kind].append(row)
    return res


def action_summary(data: dict, team_k: str, action: str) -> str:
    teams = dict(all_selected_teams(data))
    team_name = teams.get(team_k, team_k)
    groups = participants_by_team(data, team_k)
    title = "تم وضع منتخب مغادر" if action == "depart" else "تمت استعادة منتخب"
    lines = [f"✅ {title}: {team_name}", ""]
    for kind in ("abokhaled", "aboyaser"):
        rows = groups[kind]
        lines.append(f"{CONTEST_NAMES[kind]}: {len(rows)}")
        if rows:
            for r in rows:
                extra = f" — {r.get('player','')}" if kind == "aboyaser" and r.get("player") else ""
                lines.append(f"- {r.get('name','')} ({canonical_team(r.get('team',''))}{extra})")
        else:
            lines.append("- لا يوجد")
        lines.append("")
    return "\n".join(lines).strip()

# ---------- الصور ----------

def font_path() -> str:
    candidates = [
        "/usr/share/fonts/truetype/noto/NotoKufiArabic-Bold.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Bold.ttf",
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return ""


def get_font(size: int):
    if not ImageFont:
        return None
    fp = font_path()
    if fp:
        try:
            return ImageFont.truetype(fp, size)
        except Exception:
            pass
    return ImageFont.load_default()


def text_size(draw, text: str, font):
    try:
        box = draw.textbbox((0, 0), text, font=font)
        return box[2] - box[0], box[3] - box[1]
    except Exception:
        return draw.textsize(text, font=font)


def draw_text(draw, xy, text, font, fill=(255,255,255), anchor="mm"):
    draw.text(xy, ar(text), font=font, fill=fill, anchor=anchor)


def draw_round(draw, box, radius=22, fill=None, outline=None, width=1):
    try:
        draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)
    except Exception:
        draw.rectangle(box, fill=fill, outline=outline, width=width)


def render_contest_image(data: dict, kind: str) -> Path:
    if Image is None:
        raise RuntimeError("Pillow غير مثبت. ثبّت requirements.txt")
    kind = "aboyaser" if kind == "aboyaser" else "abokhaled"
    sig = data_hash({"kind": kind, "rows": contest_rows(data, kind), "departed": sorted(data.get("departed_teams", []))})
    out_path = GENERATED_DIR / f"contest_{kind}_{sig}.jpg"
    if out_path.exists() and out_path.stat().st_size > 0:
        return out_path

    rows = contest_rows(data, kind)
    width = 1200
    row_h = 64 if kind == "abokhaled" else 68
    header_h = 255
    footer_h = 90
    height = max(1500, header_h + 58 + len(rows) * row_h + footer_h)

    img = Image.new("RGB", (width, height), (6, 18, 45))
    draw = ImageDraw.Draw(img)

    # خلفية بسيطة ثابتة وسريعة
    for y in range(height):
        shade = int(18 + (y / max(1, height)) * 25)
        draw.line([(0, y), (width, y)], fill=(5, shade, 70))
    draw.ellipse((-250, -260, 430, 420), fill=(10, 65, 130))
    draw.ellipse((width-360, 30, width+230, 600), fill=(0, 42, 100))

    f_title = get_font(58)
    f_sub = get_font(30)
    f_head = get_font(25)
    f_row = get_font(28)
    f_small = get_font(22)

    title = CONTEST_NAMES[kind]
    draw_text(draw, (width//2, 72), title, f_title, fill=(255, 255, 255))
    draw_text(draw, (width//2, 132), "مسابقة المصيف 2026 — يدوي", f_sub, fill=(252, 211, 77))
    draw_text(draw, (width//2, 182), f"آخر تحديث: {now_riyadh()}", f_small, fill=(190, 226, 255))

    departed_count = sum(1 for r in rows if is_departed(data, r.get("team")))
    draw_text(draw, (width//2, 220), f"المشاركون: {len(rows)} | المستمرون: {len(rows)-departed_count} | غادروا: {departed_count}", f_small, fill=(220, 245, 255))

    x0, x1 = 65, width - 65
    y = header_h
    draw_round(draw, (x0, y, x1, y+48), radius=18, fill=(8, 45, 105), outline=(54, 169, 245), width=2)

    if kind == "abokhaled":
        draw_text(draw, (935, y+25), "المشارك", f_head)
        draw_text(draw, (585, y+25), "المنتخب", f_head)
        draw_text(draw, (245, y+25), "الحالة", f_head)
    else:
        draw_text(draw, (965, y+25), "المشارك", f_head)
        draw_text(draw, (675, y+25), "المنتخب", f_head)
        draw_text(draw, (405, y+25), "اللاعب", f_head)
        draw_text(draw, (165, y+25), "الحالة", f_head)

    y += 62
    for i, row in enumerate(rows, start=1):
        team = canonical_team(row.get("team", ""))
        out = is_departed(data, team)
        row_fill = (10, 35, 78) if not out else (72, 18, 30)
        outline = (25, 112, 190) if not out else (239, 68, 68)
        draw_round(draw, (x0, y, x1, y + row_h - 10), radius=16, fill=row_fill, outline=outline, width=1)
        cy = y + (row_h - 10) // 2
        status = "غادر" if out else "مستمر"
        status_color = (255, 120, 120) if out else (160, 255, 210)
        name = row.get("name", "")
        if kind == "abokhaled":
            draw_text(draw, (935, cy), name, f_row, fill=(255,255,255))
            draw_text(draw, (585, cy), team, f_row, fill=(253, 230, 138))
            draw_text(draw, (245, cy), status, f_row, fill=status_color)
        else:
            draw_text(draw, (965, cy), name, f_small if len(str(name)) > 12 else f_row, fill=(255,255,255))
            draw_text(draw, (675, cy), team, f_small if len(str(team)) > 10 else f_row, fill=(253, 230, 138))
            player = row.get("player", "")
            draw_text(draw, (405, cy), player, f_small if len(str(player)) > 10 else f_row, fill=(224, 242, 254))
            draw_text(draw, (165, cy), status, f_small if len(status) > 5 else f_row, fill=status_color)
        if out:
            draw.line((x0+45, cy, x1-45, cy), fill=(255, 80, 80), width=2)
        y += row_h

    draw_text(draw, (width//2, height-48), "استراحة المصيف", get_font(30), fill=(252, 211, 77))
    img.save(out_path, quality=90, optimize=True)
    return out_path

# ---------- الكيبورد ----------

def reply_keyboard():
    return ReplyKeyboardMarkup([["🏆 مسابقات المصيف"]], resize_keyboard=True, one_time_keyboard=False)


def contests_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏆 مسابقة أبوخالد", callback_data="contest|abokhaled")],
        [InlineKeyboardButton("🏆 مسابقة أبوياسر", callback_data="contest|aboyaser")],
        [InlineKeyboardButton("🛠️ إدارة المتسابقين", callback_data="manage|open")],
    ])


def manage_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🚪 منتخب غادر", callback_data="manage|depart")],
        [InlineKeyboardButton("♻️ استعادة منتخب", callback_data="manage|restore")],
        [InlineKeyboardButton("➕ إضافة متسابق", callback_data="manage|add")],
        [InlineKeyboardButton("🗑️ حذف متسابق", callback_data="manage|delete")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="menu|contests")],
    ])


def choose_contest_keyboard(prefix: str):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏆 أبوخالد", callback_data=f"{prefix}|abokhaled")],
        [InlineKeyboardButton("🏆 أبوياسر", callback_data=f"{prefix}|aboyaser")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="manage|home")],
    ])


def teams_keyboard(data: dict, mode: str):
    only_departed = mode == "restore"
    teams = all_selected_teams(data, only_departed=only_departed)
    if not teams:
        return InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ رجوع", callback_data="manage|home")]])
    rows = []
    line = []
    dep = departed_keys(data)
    for idx, (k, name) in enumerate(teams):
        status = "🔴 مغادر" if k in dep else "🟢 مستمر"
        label = f"{name} {status}"
        line.append(InlineKeyboardButton(label[:34], callback_data=f"team|{mode}|{idx}"))
        if len(line) == 2:
            rows.append(line)
            line = []
    if line:
        rows.append(line)
    rows.append([InlineKeyboardButton("⬅️ رجوع", callback_data="manage|home")])
    return InlineKeyboardMarkup(rows)


def contestants_keyboard(data: dict, kind: str):
    rows = contest_rows(data, kind)
    kb = []
    for idx, r in enumerate(rows):
        label = f"{r.get('name','')} — {canonical_team(r.get('team',''))}"[:38]
        kb.append([InlineKeyboardButton(label, callback_data=f"delrow|{kind}|{idx}")])
    kb.append([InlineKeyboardButton("⬅️ رجوع", callback_data="manage|delete")])
    return InlineKeyboardMarkup(kb)

# ---------- الأوامر والردود ----------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.effective_message.reply_text(
        "حياك في بوت مسابقات المصيف اليدوي 🏆\nاختر من القائمة:",
        reply_markup=reply_keyboard(),
    )


async def show_contests_menu(message):
    await message.reply_text("🏆 مسابقة المصيف يدوي\nاختر:", reply_markup=contests_keyboard())


async def send_contest(message, kind: str):
    data = load_data()
    title = CONTEST_NAMES.get(kind, "المسابقة")
    wait = await message.reply_text(f"⏳ جاري تجهيز صورة {title}...")
    try:
        path = render_contest_image(data, kind)
        try:
            await wait.delete()
        except Exception:
            pass
        with path.open("rb") as f:
            await message.reply_photo(photo=f, caption=f"🏆 {title}")
    except Exception as e:
        try:
            await wait.edit_text(f"تعذر تجهيز الصورة: {e}")
        except Exception:
            await message.reply_text(f"تعذر تجهيز الصورة: {e}")


async def callbacks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    if not q:
        return
    data_cb = q.data or ""
    try:
        await q.answer()
    except Exception:
        pass
    data = load_data()

    if data_cb == "menu|contests":
        context.user_data.pop("flow", None)
        await q.message.reply_text("🏆 مسابقة المصيف يدوي\nاختر:", reply_markup=contests_keyboard())
        return

    if data_cb.startswith("contest|"):
        kind = data_cb.split("|", 1)[1]
        await send_contest(q.message, kind)
        return

    if data_cb == "manage|open":
        context.user_data["flow"] = {"type": "password"}
        await q.message.reply_text("اكتب الرقم السري:")
        return

    if data_cb == "manage|home":
        if not context.user_data.get("manager_authed"):
            context.user_data["flow"] = {"type": "password"}
            await q.message.reply_text("اكتب الرقم السري:")
            return
        await q.message.reply_text("🛠️ إدارة المتسابقين", reply_markup=manage_keyboard())
        return

    if data_cb == "manage|depart":
        if not context.user_data.get("manager_authed"):
            context.user_data["flow"] = {"type": "password"}
            await q.message.reply_text("اكتب الرقم السري:")
            return
        await q.message.reply_text("🚪 اختر المنتخب اللي غادر:", reply_markup=teams_keyboard(data, "depart"))
        return

    if data_cb == "manage|restore":
        if not context.user_data.get("manager_authed"):
            context.user_data["flow"] = {"type": "password"}
            await q.message.reply_text("اكتب الرقم السري:")
            return
        await q.message.reply_text("♻️ اختر المنتخب اللي تبي تستعيده:", reply_markup=teams_keyboard(data, "restore"))
        return

    if data_cb.startswith("team|"):
        if not context.user_data.get("manager_authed"):
            context.user_data["flow"] = {"type": "password"}
            await q.message.reply_text("اكتب الرقم السري:")
            return
        _, mode, idx_s = data_cb.split("|", 2)
        teams = all_selected_teams(data, only_departed=(mode == "restore"))
        try:
            idx = int(idx_s)
            team_k, team_name = teams[idx]
        except Exception:
            await q.message.reply_text("الاختيار غير صحيح.", reply_markup=manage_keyboard())
            return
        dep = set(data.get("departed_teams", []))
        if mode == "depart":
            dep.add(team_k)
            data["departed_teams"] = sorted(dep)
            save_data(data)
            await q.message.reply_text(action_summary(data, team_k, "depart"), reply_markup=manage_keyboard())
        else:
            dep.discard(team_k)
            data["departed_teams"] = sorted(dep)
            save_data(data)
            await q.message.reply_text(action_summary(data, team_k, "restore"), reply_markup=manage_keyboard())
        return

    if data_cb == "manage|add":
        if not context.user_data.get("manager_authed"):
            context.user_data["flow"] = {"type": "password"}
            await q.message.reply_text("اكتب الرقم السري:")
            return
        await q.message.reply_text("اختر المسابقة للإضافة:", reply_markup=choose_contest_keyboard("addcontest"))
        return

    if data_cb.startswith("addcontest|"):
        kind = data_cb.split("|", 1)[1]
        context.user_data["flow"] = {"type": "add", "kind": kind, "step": "name"}
        await q.message.reply_text("اكتب اسم المتسابق:")
        return

    if data_cb == "manage|delete":
        if not context.user_data.get("manager_authed"):
            context.user_data["flow"] = {"type": "password"}
            await q.message.reply_text("اكتب الرقم السري:")
            return
        await q.message.reply_text("اختر المسابقة للحذف:", reply_markup=choose_contest_keyboard("delcontest"))
        return

    if data_cb.startswith("delcontest|"):
        kind = data_cb.split("|", 1)[1]
        await q.message.reply_text(f"اختر المتسابق من {CONTEST_NAMES[kind]}:", reply_markup=contestants_keyboard(data, kind))
        return

    if data_cb.startswith("delrow|"):
        _, kind, idx_s = data_cb.split("|", 2)
        try:
            idx = int(idx_s)
            row = contest_rows(data, kind)[idx]
        except Exception:
            await q.message.reply_text("الاختيار غير صحيح.", reply_markup=manage_keyboard())
            return
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ تأكيد الحذف", callback_data=f"delconfirm|{kind}|{idx}")],
            [InlineKeyboardButton("❌ إلغاء", callback_data=f"delcontest|{kind}")],
        ])
        await q.message.reply_text(f"تأكيد حذف:\n{row.get('name','')} — {canonical_team(row.get('team',''))}", reply_markup=kb)
        return

    if data_cb.startswith("delconfirm|"):
        _, kind, idx_s = data_cb.split("|", 2)
        try:
            idx = int(idx_s)
            row = data["contests"][kind].pop(idx)
            save_data(data)
            await q.message.reply_text(f"✅ تم حذف {row.get('name','')} من {CONTEST_NAMES[kind]}", reply_markup=manage_keyboard())
        except Exception:
            await q.message.reply_text("تعذر الحذف، جرّب مرة ثانية.", reply_markup=manage_keyboard())
        return

    await q.message.reply_text("اختر من القائمة:", reply_markup=contests_keyboard())


async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.effective_message.text or "").strip()
    n = normalize_ar(text)

    if text == "🏆 مسابقات المصيف" or n in {"مسابقات المصيف", "المسابقات", "مسابقه المصيف"}:
        await show_contests_menu(update.effective_message)
        return

    flow = context.user_data.get("flow") or {}
    ftype = flow.get("type")

    if ftype == "password":
        if normalize_ar(text) == normalize_ar(PASSWORD):
            context.user_data["manager_authed"] = True
            context.user_data.pop("flow", None)
            await update.effective_message.reply_text("✅ تم الدخول لإدارة المتسابقين", reply_markup=manage_keyboard())
        else:
            await update.effective_message.reply_text("❌ الرقم السري خطأ. اكتب الرقم السري مرة ثانية:")
        return

    if ftype == "add":
        kind = flow.get("kind")
        step = flow.get("step")
        if kind not in CONTEST_NAMES:
            context.user_data.pop("flow", None)
            await update.effective_message.reply_text("انتهت العملية. ارجع للإدارة.", reply_markup=manage_keyboard())
            return
        if step == "name":
            flow["name"] = text
            flow["step"] = "team"
            context.user_data["flow"] = flow
            await update.effective_message.reply_text("اكتب المنتخب:")
            return
        if step == "team":
            flow["team"] = canonical_team(text)
            if kind == "aboyaser":
                flow["step"] = "player"
                context.user_data["flow"] = flow
                await update.effective_message.reply_text("اكتب اللاعب:")
                return
            data = load_data()
            data["contests"][kind].append({"name": flow["name"], "team": flow["team"]})
            save_data(data)
            context.user_data.pop("flow", None)
            await update.effective_message.reply_text(
                f"✅ تم إضافة {flow['name']} إلى {CONTEST_NAMES[kind]}\nالمنتخب: {flow['team']}",
                reply_markup=manage_keyboard(),
            )
            return
        if step == "player":
            data = load_data()
            data["contests"][kind].append({"name": flow["name"], "team": flow["team"], "player": text})
            save_data(data)
            context.user_data.pop("flow", None)
            await update.effective_message.reply_text(
                f"✅ تم إضافة {flow['name']} إلى {CONTEST_NAMES[kind]}\nالمنتخب: {flow['team']}\nاللاعب: {text}",
                reply_markup=manage_keyboard(),
            )
            return

    # أوامر نصية احتياطية سريعة، لكنها ليست مطلوبة للواجهة.
    if n.startswith("غادر ") or n.startswith("غادرو "):
        team = text.split(maxsplit=1)[1] if len(text.split(maxsplit=1)) > 1 else ""
        data = load_data()
        tk = team_key(team)
        dep = set(data.get("departed_teams", []))
        dep.add(tk)
        data["departed_teams"] = sorted(dep)
        save_data(data)
        await update.effective_message.reply_text(action_summary(data, tk, "depart"), reply_markup=manage_keyboard())
        return

    if n.startswith("رجع ") or n.startswith("استعاده "):
        team = text.split(maxsplit=1)[1] if len(text.split(maxsplit=1)) > 1 else ""
        data = load_data()
        tk = team_key(team)
        dep = set(data.get("departed_teams", []))
        dep.discard(tk)
        data["departed_teams"] = sorted(dep)
        save_data(data)
        await update.effective_message.reply_text(action_summary(data, tk, "restore"), reply_markup=manage_keyboard())
        return

    await update.effective_message.reply_text("اختر من القائمة:", reply_markup=reply_keyboard())


def main():
    if not TOKEN:
        raise RuntimeError("ضع توكن البوت في متغير بيئة: BOT_TOKEN أو TELEGRAM_BOT_TOKEN أو TELEGRAM_TOKEN أو TOKEN")
    load_data()
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("menu", start))
    app.add_handler(CallbackQueryHandler(callbacks))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    log.info("بوت مسابقات المصيف اليدوي يعمل الآن")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
