#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
بوت مسابقة المصيف اليدوي — نسخة نظيفة.
- لا ESPN
- لا بث مباشر
- لا نتائج مباريات
- لا إشعارات
- لا سحب مواقع

يعتمد فقط على تيليجرام + ملف contest_data.json المحلي.
"""

from __future__ import annotations

import os
import re
import json
import hashlib
import asyncio
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Any, Tuple

from telegram import Update, InlineKeyboardMarkup, InlineKeyboardButton, ReplyKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackQueryHandler, MessageHandler, ContextTypes, filters

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

BOT_TOKEN = os.environ.get("BOT_TOKEN") or os.environ.get("TELEGRAM_BOT_TOKEN") or os.environ.get("TOKEN") or ""
DATA_DIR = Path(os.environ.get("DATA_DIR", "."))
DATA_FILE = Path(os.environ.get("CONTEST_DATA_FILE", str(DATA_DIR / "contest_data.json")))
GENERATED_DIR = Path(os.environ.get("GENERATED_DIR", "generated"))
SECRET = "الملحق"

# البيانات الأساسية مأخوذة من نسخة المسابقة السابقة المعتمدة.
DEFAULT_ABOKHALED = [{'name': 'أبو خالد', 'team': 'إنجلترا'}, {'name': 'سلمان أحمد', 'team': 'إسبانيا'}, {'name': 'أبوداحم', 'team': 'البرازيل'}, {'name': 'نواف فارس', 'team': 'فرنسا'}, {'name': 'أبو نايف', 'team': 'البرتغال'}, {'name': 'أبو راكان', 'team': 'هولندا'}, {'name': 'مشعل', 'team': 'البرتغال'}, {'name': 'نايف', 'team': 'هولندا'}, {'name': 'محمد عبدالرحمن', 'team': 'البرتغال'}, {'name': 'سلطان', 'team': 'فرنسا'}, {'name': 'خالد', 'team': 'إسبانيا'}, {'name': 'عادل', 'team': 'البرتغال'}, {'name': 'أبو عبدالله', 'team': 'السعودية'}, {'name': 'فهد فارس', 'team': 'ألمانيا'}, {'name': 'زياد', 'team': 'السعودية'}, {'name': 'فارس سالم', 'team': 'فرنسا'}, {'name': 'مشاري عبدالعزيز', 'team': 'إسبانيا'}, {'name': 'طلال عبدالله', 'team': 'الأرجنتين'}, {'name': 'عبدالعزيز', 'team': 'البرتغال'}, {'name': 'أبوتركي', 'team': 'نيوزيلندا'}, {'name': 'أبوفارس', 'team': 'إسبانيا'}, {'name': 'عبدالله', 'team': 'فرنسا'}, {'name': 'أبو طلال', 'team': 'ساحل العاج'}, {'name': 'خالد', 'team': 'إسبانيا'}, {'name': 'يزيد', 'team': 'إنجلترا'}, {'name': 'أبو يزيد', 'team': 'إنجلترا'}]
DEFAULT_ABOYASER = [{'name': 'أبو فارس', 'team': 'البرازيل', 'player': 'ديمبلي'}, {'name': 'أبو هوى', 'team': 'فرنسا', 'player': 'يامال'}, {'name': 'تركي محسن', 'team': 'البرازيل', 'player': 'يامال'}, {'name': 'زيكا', 'team': 'البرتغال', 'player': 'مبابي'}, {'name': 'طلال عبدالله', 'team': 'إنجلترا', 'player': 'يامال'}, {'name': 'عبدالله إبراهيم', 'team': 'فرنسا', 'player': 'مبابي'}, {'name': 'بدران', 'team': 'فرنسا', 'player': 'أوليسي'}, {'name': 'فهد فارس', 'team': 'ألمانيا', 'player': 'يامال'}, {'name': 'أبو صنت', 'team': 'السعودية', 'player': 'سعود عبدالحميد'}, {'name': 'يزيد', 'team': 'إسبانيا', 'player': 'مبابي'}, {'name': 'سلطان أحمد', 'team': 'فرنسا', 'player': 'مبابي'}, {'name': 'نواف فارس', 'team': 'إنجلترا', 'player': 'هاري كين'}, {'name': 'الأمير', 'team': 'إسبانيا', 'player': 'يامال'}, {'name': 'فارس سالم', 'team': 'فرنسا', 'player': 'يامال'}, {'name': 'مشاري عبدالعزيز', 'team': 'إسبانيا', 'player': 'يامال'}, {'name': 'هندسة', 'team': 'البرازيل', 'player': 'يامال'}, {'name': 'سلمان أحمد', 'team': 'إسبانيا', 'player': 'يامال'}, {'name': 'نايف حمود', 'team': 'الأرجنتين', 'player': 'أبو محمد'}, {'name': 'عبدالرحمن سالم', 'team': 'البرازيل', 'player': 'مبابي'}, {'name': 'محمد محسن', 'team': 'البرازيل', 'player': 'بيدرو'}, {'name': 'أبو شنب', 'team': 'فرنسا', 'player': 'مبابي'}, {'name': 'ممدوح غزاي', 'team': 'إسبانيا', 'player': 'فيرتز'}, {'name': 'جلعده', 'team': 'إسبانيا', 'player': 'مبابي'}, {'name': 'خالد عبدالرحمن', 'team': 'إسبانيا', 'player': 'يامال'}, {'name': 'محمد عبدالرحمن', 'team': 'البرتغال', 'player': 'مبابي'}, {'name': 'سلطان رباح', 'team': 'إسبانيا', 'player': 'يامال'}, {'name': 'عادل', 'team': 'البرتغال', 'player': 'كين'}]

TEAM_ALIASES = {
    "البرازيد": "البرازيل",
    "برازيد": "البرازيل",
    "برازيل": "البرازيل",
    "المانيا": "ألمانيا",
    "ألمانيا": "ألمانيا",
    "الارجنتين": "الأرجنتين",
    "ارجنتين": "الأرجنتين",
    "هولندا": "هولندا",
    "المغرب": "المغرب",
    "اسبانيا": "إسبانيا",
    "إسبانيا": "إسبانيا",
    "فرنسا": "فرنسا",
    "البرتغال": "البرتغال",
    "انجلترا": "إنجلترا",
    "إنجلترا": "إنجلترا",
    "السعودية": "السعودية",
    "نيوزيلندا": "نيوزيلندا",
    "ساحل العاج": "ساحل العاج",
}


def normalize_name(text: Any) -> str:
    s = "" if text is None else str(text)
    s = s.strip()
    s = re.sub(r"[\u064B-\u065F\u0670]", "", s)
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ى", "ي")
    s = s.replace("ة", "ه")
    s = re.sub(r"\s+", " ", s)
    return s.strip().lower()


def canonical_team(team: Any) -> str:
    raw = "" if team is None else str(team).strip()
    if not raw:
        return ""
    key = normalize_name(raw)
    for a, b in TEAM_ALIASES.items():
        if normalize_name(a) == key:
            return b
    # طابق المنتخبات الموجودة في البيانات الحالية عشان نحافظ على نفس الصيغة.
    try:
        data = load_data()
        for row in data.get("abokhaled", []) + data.get("aboyaser", []):
            t = row.get("team", "")
            if normalize_name(t) == key:
                return str(t).strip()
    except Exception:
        pass
    return raw


def team_key(team: Any) -> str:
    return normalize_name(canonical_team(team))


def ar_text(text: Any) -> str:
    text = "" if text is None else str(text)
    if arabic_reshaper and get_display:
        try:
            return get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
    return text


def has_arabic(text: Any) -> bool:
    return bool(re.search(r"[\u0600-\u06FF]", str(text or "")))


def font_candidates() -> List[str]:
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
        "DejaVuSans.ttf",
    ]
    bases = ["", ".", "/app", "/mnt/data", "/usr/share/fonts/truetype/noto", "/usr/share/fonts/truetype/dejavu", "/usr/share/fonts"]
    paths: List[str] = []
    for base in bases:
        for name in names:
            paths.append(os.path.join(base, name) if base else name)
    return paths


def get_font(size: int):
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


def font_size(font, default=24) -> int:
    try:
        return int(getattr(font, "size", default) or default)
    except Exception:
        return default


def get_arabic_fallback_font(size: int):
    return get_font(size)


def clean_draw_text(text: Any) -> str:
    s = "" if text is None else str(text)
    # نحذف الإيموجي فقط من الصورة حتى لا تتحول مربعات في الخطوط.
    replacements = {
        "🏆": "",
        "👑": "",
        "🧤": "",
        "👥": "",
        "⚽": "",
        "🔥": "",
        "😅": "",
        "⚔️": "",
        "⚔": "",
        "\ufe0f": "",
        "\u200f": "",
        "\u200e": "",
        "\u2066": "",
        "\u2067": "",
        "\u2068": "",
        "\u2069": "",
    }
    for a, b in replacements.items():
        s = s.replace(a, b)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def text_width(draw, text, font) -> int:
    text = clean_draw_text(text)
    try:
        if has_arabic(text) and PIL_RAQM:
            bbox = draw.textbbox((0, 0), text, font=font, direction="rtl", language="ar")
        elif has_arabic(text):
            f = get_arabic_fallback_font(font_size(font, 24))
            bbox = draw.textbbox((0, 0), ar_text(text), font=f)
        else:
            bbox = draw.textbbox((0, 0), text, font=font)
        return int(bbox[2] - bbox[0])
    except Exception:
        return len(str(text)) * max(10, int(font_size(font, 24) * 0.55))


def wrap_text(draw, text, font, max_width: int) -> List[str]:
    text = clean_draw_text(text)
    final_lines: List[str] = []
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


def draw_text(draw, xy, text, font, fill="white", anchor="mm", align="center", max_width=None, spacing=8):
    text = clean_draw_text(text)
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


def rounded_rect(draw, box, radius=24, fill=None, outline=None, width=1):
    try:
        draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)
    except Exception:
        draw.rectangle(box, fill=fill, outline=outline, width=width)


def now_riyadh_text() -> str:
    dt = datetime.utcnow() + timedelta(hours=3)
    suffix = "صباحًا" if dt.hour < 12 else "مساءً"
    return dt.strftime("%d %m %Y - %I:%M") + f" {suffix}"


def contest_bg(width: int, height: int):
    if not Image:
        raise RuntimeError("Pillow غير مثبت")
    img = Image.new("RGB", (width, height), "#06152F")
    draw = ImageDraw.Draw(img)
    for y in range(height):
        t = y / max(1, height)
        draw.line((0, y, width, y), fill=(int(3 + 4 * t), int(14 + 20 * t), int(45 + 55 * t)))
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.ellipse((120, -230, width - 80, 470), fill=(15, 92, 190, 95))
    od.ellipse((width * 0.55, 420, width * 1.15, 1030), fill=(0, 102, 230, 50))
    od.ellipse((-220, height * 0.55, width * 0.45, height + 180), fill=(0, 115, 255, 34))
    od.rectangle((width * 0.66, 520, width * 0.88, 720), fill=(0, 115, 255, 38))
    od.rectangle((width * 0.90, 610, width * 1.08, 860), fill=(0, 105, 230, 42))
    try:
        overlay = overlay.filter(ImageFilter.GaussianBlur(8))
    except Exception:
        pass
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    return img, ImageDraw.Draw(img)


def default_data() -> Dict[str, Any]:
    return {
        "version": 1,
        "out_teams": [],
        "abokhaled": DEFAULT_ABOKHALED,
        "aboyaser": DEFAULT_ABOYASER,
    }


def load_data() -> Dict[str, Any]:
    if DATA_FILE.exists():
        try:
            with DATA_FILE.open("r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                # تأكد من وجود المفاتيح بعد أي تحديث.
                data.setdefault("out_teams", [])
                data.setdefault("abokhaled", DEFAULT_ABOKHALED)
                data.setdefault("aboyaser", DEFAULT_ABOYASER)
                return data
        except Exception:
            pass
    data = default_data()
    save_data(data)
    return data


def save_data(data: Dict[str, Any]) -> None:
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    with DATA_FILE.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_rows(kind: str) -> List[Dict[str, Any]]:
    data = load_data()
    return list(data.get(kind, []))


def is_team_out(team: str) -> bool:
    data = load_data()
    out = set(team_key(x) for x in data.get("out_teams", []))
    return team_key(team) in out


def set_team_out(team: str, out: bool) -> Tuple[str, Dict[str, List[str]]]:
    data = load_data()
    team = canonical_team(team)
    key = team_key(team)
    current = list(data.get("out_teams", []))
    current_keys = [team_key(x) for x in current]
    if out:
        if key not in current_keys:
            current.append(team)
    else:
        current = [x for x in current if team_key(x) != key]
    data["out_teams"] = current
    save_data(data)
    return team, affected_by_team(team)


def affected_by_team(team: str) -> Dict[str, List[str]]:
    data = load_data()
    key = team_key(team)
    res = {"abokhaled": [], "aboyaser": []}
    for kind in ["abokhaled", "aboyaser"]:
        for row in data.get(kind, []):
            if team_key(row.get("team")) == key:
                res[kind].append(str(row.get("name", "")))
    return res


def unique_teams(mode: str = "all") -> List[str]:
    data = load_data()
    teams: List[str] = []
    seen = set()
    out_keys = set(team_key(x) for x in data.get("out_teams", []))
    for row in data.get("abokhaled", []) + data.get("aboyaser", []):
        team = canonical_team(row.get("team", ""))
        if not team:
            continue
        k = team_key(team)
        if k in seen:
            continue
        if mode == "restore" and k not in out_keys:
            continue
        teams.append(team)
        seen.add(k)
    # لو فيه منتخب مغادر وما عاد له مشاركين، نخليه في الاستعادة احتياطًا.
    if mode == "restore":
        for team in data.get("out_teams", []):
            k = team_key(team)
            if k not in seen:
                teams.append(canonical_team(team))
                seen.add(k)
    return teams


def contest_title(kind: str) -> str:
    return "مسابقة أبوخالد" if kind == "abokhaled" else "مسابقة أبوياسر"


def render_contest(kind: str = "abokhaled") -> str:
    data = load_data()
    rows = data.get(kind, [])
    title = contest_title(kind)
    subtitle = "ترشيحات الفوز بكأس العالم" if kind == "abokhaled" else "المشارك / المنتخب / اللاعب"
    width = 1080
    row_h = 58 if kind == "abokhaled" else 62
    height = max(1500, 260 + len(rows) * row_h + 150)
    img, draw = contest_bg(width, height)
    draw_text(draw, (width // 2, 70), title, get_font(56), fill="#FFFFFF", max_width=900)
    draw_text(draw, (width // 2, 132), subtitle, get_font(32), fill="#FDE68A", max_width=880)
    draw_text(draw, (width // 2, 178), "آخر تحديث: الآن", get_font(24), fill="#CFE8FF", max_width=820)

    x0, x1 = 60, width - 60
    y = 225
    rounded_rect(draw, (x0, y, x1, y + 46), radius=18, fill="#0B2A5CCC", outline="#38BDF8", width=2)
    if kind == "abokhaled":
        draw_text(draw, (855, y + 24), "المشارك", get_font(24), fill="#FFFFFF")
        draw_text(draw, (520, y + 24), "المنتخب", get_font(24), fill="#FFFFFF")
        draw_text(draw, (215, y + 24), "الحالة", get_font(24), fill="#FFFFFF")
    else:
        draw_text(draw, (875, y + 24), "المشارك", get_font(23), fill="#FFFFFF")
        draw_text(draw, (575, y + 24), "المنتخب", get_font(23), fill="#FFFFFF")
        draw_text(draw, (300, y + 24), "اللاعب", get_font(23), fill="#FFFFFF")
        draw_text(draw, (115, y + 24), "الحالة", get_font(23), fill="#FFFFFF")
    y += 56

    for r in rows:
        team = canonical_team(r.get("team", ""))
        out = is_team_out(team)
        fill = "#071A36DD" if not out else "#2B1018DD"
        outline = "#1D9BFF88" if not out else "#FF4B4B99"
        rounded_rect(draw, (x0, y, x1, y + row_h - 8), radius=16, fill=fill, outline=outline, width=1)
        cy = y + (row_h - 8) // 2
        name = str(r.get("name", ""))
        status = "غادر" if out else "مستمر"
        status_color = "#FF5555" if out else "#A7F3D0"
        if kind == "abokhaled":
            draw_text(draw, (855, cy), name, get_font(26), fill="#FFFFFF", max_width=380)
            draw_text(draw, (520, cy), team, get_font(26), fill="#FDE68A", max_width=270)
            draw_text(draw, (215, cy), status, get_font(24), fill=status_color, max_width=220)
        else:
            draw_text(draw, (875, cy), name, get_font(23), fill="#FFFFFF", max_width=330)
            draw_text(draw, (575, cy), team, get_font(23), fill="#FDE68A", max_width=235)
            draw_text(draw, (300, cy), str(r.get("player", "")), get_font(23), fill="#E0F2FE", max_width=220)
            draw_text(draw, (115, cy), status, get_font(22), fill=status_color, max_width=110)
        if out:
            try:
                draw.line((x0 + 55, cy, x1 - 55, cy), fill="#FF5A5A", width=2)
            except Exception:
                pass
        y += row_h

    draw_text(draw, (width // 2, height - 60), "مونديال المصيف 2026", get_font(28), fill="#FBBF24")
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    sig = hashlib.md5(json.dumps(data, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:10]
    out_path = GENERATED_DIR / f"contest_{kind}_{sig}.jpg"
    img.save(out_path, quality=94, optimize=True)
    return str(out_path)


def main_reply_keyboard():
    return ReplyKeyboardMarkup([["🏆 مسابقات المصيف"]], resize_keyboard=True, one_time_keyboard=False)


def contests_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏆 مسابقة أبوخالد", callback_data="show|abokhaled")],
        [InlineKeyboardButton("🏆 مسابقة أبوياسر", callback_data="show|aboyaser")],
        [InlineKeyboardButton("🛠️ إدارة المتسابقين", callback_data="manage")],
    ])


def admin_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🚪 منتخب غادر", callback_data="pick|out")],
        [InlineKeyboardButton("♻️ استعادة منتخب", callback_data="pick|restore")],
        [InlineKeyboardButton("➕ إضافة متسابق", callback_data="add_menu")],
        [InlineKeyboardButton("🗑️ حذف متسابق", callback_data="delete_menu")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="menu")],
    ])


def kind_keyboard(prefix: str):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏆 أبوخالد", callback_data=f"{prefix}|abokhaled")],
        [InlineKeyboardButton("🏆 أبوياسر", callback_data=f"{prefix}|aboyaser")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="admin_menu")],
    ])


def teams_keyboard(mode: str):
    teams = unique_teams("restore" if mode == "restore" else "all")
    rows = []
    line = []
    for idx, team in enumerate(teams):
        status = "🔴 مغادر" if is_team_out(team) else "🟢 مستمر"
        label = f"{team} {status}"
        line.append(InlineKeyboardButton(label[:45], callback_data=f"team|{mode}|{idx}"))
        if len(line) == 2:
            rows.append(line)
            line = []
    if line:
        rows.append(line)
    if not rows:
        rows.append([InlineKeyboardButton("لا يوجد منتخبات", callback_data="admin_menu")])
    rows.append([InlineKeyboardButton("⬅️ رجوع", callback_data="admin_menu")])
    return InlineKeyboardMarkup(rows)


def participants_keyboard(kind: str):
    rows = []
    data = load_data()
    items = data.get(kind, [])
    for idx, row in enumerate(items):
        label = f"{row.get('name','')} — {row.get('team','')}"
        rows.append([InlineKeyboardButton(label[:55], callback_data=f"del_confirm|{kind}|{idx}")])
    if not rows:
        rows.append([InlineKeyboardButton("لا يوجد متسابقين", callback_data="admin_menu")])
    rows.append([InlineKeyboardButton("⬅️ رجوع", callback_data="admin_menu")])
    return InlineKeyboardMarkup(rows)


async def send_photo_path(message, path: str, caption: str = ""):
    with open(path, "rb") as f:
        await message.reply_photo(photo=f, caption=caption)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.effective_message.reply_text("حياك الله 👋\nاختر من القائمة:", reply_markup=main_reply_keyboard())


async def show_contests_menu(message):
    await message.reply_text("🏆 مسابقات المصيف\n\nاختر:", reply_markup=contests_keyboard())


async def send_contest_image(message, kind: str):
    wait = await message.reply_text(f"⏳ جاري تجهيز صورة {contest_title(kind)}...")
    try:
        path = await asyncio.to_thread(render_contest, kind)
        try:
            await wait.delete()
        except Exception:
            pass
        await send_photo_path(message, path, f"🏆 {contest_title(kind)}")
    except Exception as e:
        try:
            await wait.edit_text(f"❌ تعذر تجهيز الصورة\n{str(e)[:200]}")
        except Exception:
            await message.reply_text(f"❌ تعذر تجهيز الصورة\n{str(e)[:200]}")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.effective_message
    text = (msg.text or "").strip()
    state = context.user_data.get("state")

    if state == "password":
        if normalize_name(text) == normalize_name(SECRET):
            context.user_data["authed"] = True
            context.user_data.pop("state", None)
            await msg.reply_text("✅ تم فتح إدارة المتسابقين", reply_markup=admin_keyboard())
        else:
            await msg.reply_text("❌ الرقم غير صحيح")
        return

    if state == "add_name":
        context.user_data["new_name"] = text
        context.user_data["state"] = "add_team"
        await msg.reply_text("اكتب اسم المنتخب:")
        return

    if state == "add_team":
        context.user_data["new_team"] = canonical_team(text)
        if context.user_data.get("add_kind") == "aboyaser":
            context.user_data["state"] = "add_player"
            await msg.reply_text("اكتب اسم اللاعب:")
        else:
            await finish_add(update, context)
        return

    if state == "add_player":
        context.user_data["new_player"] = text
        await finish_add(update, context)
        return

    if text in {"🏆 مسابقات المصيف", "مسابقات المصيف", "المسابقات", "مسابقة المصيف"}:
        await show_contests_menu(msg)
        return

    await msg.reply_text("اختر من القائمة:", reply_markup=main_reply_keyboard())


async def finish_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kind = context.user_data.get("add_kind")
    name = context.user_data.get("new_name", "").strip()
    team = canonical_team(context.user_data.get("new_team", ""))
    player = context.user_data.get("new_player", "").strip()
    if not kind or not name or not team:
        context.user_data.clear()
        await update.effective_message.reply_text("❌ نقصت بيانات الإضافة. أعد المحاولة.", reply_markup=admin_keyboard())
        return
    data = load_data()
    row = {"name": name, "team": team}
    if kind == "aboyaser":
        row["player"] = player
    data.setdefault(kind, []).append(row)
    save_data(data)
    context.user_data.clear()
    context.user_data["authed"] = True
    await update.effective_message.reply_text(f"✅ تم إضافة {name} في {contest_title(kind)}", reply_markup=admin_keyboard())


async def callbacks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    if not q:
        return
    data = q.data or ""
    try:
        await q.answer()
    except Exception:
        pass

    if data == "menu":
        await q.message.reply_text("🏆 مسابقات المصيف\n\nاختر:", reply_markup=contests_keyboard())
        return

    if data.startswith("show|"):
        kind = data.split("|", 1)[1]
        await send_contest_image(q.message, kind)
        return

    if data == "manage":
        if context.user_data.get("authed"):
            await q.message.reply_text("🛠️ إدارة المتسابقين", reply_markup=admin_keyboard())
        else:
            context.user_data["state"] = "password"
            await q.message.reply_text("🔐 اكتب الرقم السري:")
        return

    if data == "admin_menu":
        if not context.user_data.get("authed"):
            context.user_data["state"] = "password"
            await q.message.reply_text("🔐 اكتب الرقم السري:")
            return
        await q.message.reply_text("🛠️ إدارة المتسابقين", reply_markup=admin_keyboard())
        return

    if not context.user_data.get("authed"):
        context.user_data["state"] = "password"
        await q.message.reply_text("🔐 اكتب الرقم السري:")
        return

    if data.startswith("pick|"):
        mode = data.split("|", 1)[1]
        title = "اختر المنتخب الذي غادر:" if mode == "out" else "اختر المنتخب الذي تريد استعادته:"
        await q.message.reply_text(title, reply_markup=teams_keyboard(mode))
        return

    if data.startswith("team|"):
        _, mode, idx_s = data.split("|")
        idx = int(idx_s)
        teams = unique_teams("restore" if mode == "restore" else "all")
        if idx < 0 or idx >= len(teams):
            await q.message.reply_text("❌ اختيار غير صحيح", reply_markup=admin_keyboard())
            return
        team = teams[idx]
        team, affected = set_team_out(team, mode == "out")
        action = "استبعاد" if mode == "out" else "استعادة"
        lines = [f"✅ تم {action} مختاري {team}", ""]
        for kind, label in [("abokhaled", "أبوخالد"), ("aboyaser", "أبوياسر")]:
            names = affected.get(kind, [])
            lines.append(f"{label}:" )
            if names:
                lines.extend([f"- {x}" for x in names])
            else:
                lines.append("- لا يوجد")
            lines.append("")
        await q.message.reply_text("\n".join(lines).strip(), reply_markup=admin_keyboard())
        return

    if data == "add_menu":
        await q.message.reply_text("اختر المسابقة للإضافة:", reply_markup=kind_keyboard("add_kind"))
        return

    if data.startswith("add_kind|"):
        kind = data.split("|", 1)[1]
        context.user_data["add_kind"] = kind
        context.user_data["state"] = "add_name"
        await q.message.reply_text("اكتب اسم المتسابق:")
        return

    if data == "delete_menu":
        await q.message.reply_text("اختر المسابقة للحذف:", reply_markup=kind_keyboard("delete_kind"))
        return

    if data.startswith("delete_kind|"):
        kind = data.split("|", 1)[1]
        await q.message.reply_text(f"اختر المتسابق من {contest_title(kind)}:", reply_markup=participants_keyboard(kind))
        return

    if data.startswith("del_confirm|"):
        _, kind, idx_s = data.split("|")
        idx = int(idx_s)
        rows = load_data().get(kind, [])
        if idx < 0 or idx >= len(rows):
            await q.message.reply_text("❌ المتسابق غير موجود", reply_markup=admin_keyboard())
            return
        row = rows[idx]
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ تأكيد الحذف", callback_data=f"del_yes|{kind}|{idx}")],
            [InlineKeyboardButton("❌ إلغاء", callback_data="admin_menu")],
        ])
        await q.message.reply_text(f"هل تريد حذف {row.get('name','')}؟", reply_markup=kb)
        return

    if data.startswith("del_yes|"):
        _, kind, idx_s = data.split("|")
        idx = int(idx_s)
        data_obj = load_data()
        rows = data_obj.get(kind, [])
        if idx < 0 or idx >= len(rows):
            await q.message.reply_text("❌ المتسابق غير موجود", reply_markup=admin_keyboard())
            return
        row = rows.pop(idx)
        data_obj[kind] = rows
        save_data(data_obj)
        await q.message.reply_text(f"✅ تم حذف {row.get('name','')} من {contest_title(kind)}", reply_markup=admin_keyboard())
        return

    await q.message.reply_text("اختر من القائمة:", reply_markup=contests_keyboard())


def main():
    if not BOT_TOKEN:
        raise RuntimeError("ضع توكن البوت في BOT_TOKEN أو TELEGRAM_BOT_TOKEN")
    load_data()
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(callbacks))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    print("Contest manual bot is running...")
    app.run_polling(close_loop=False)


if __name__ == "__main__":
    main()
