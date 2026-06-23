#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V56 patch launcher for MONDIAL AL MASEEF 2026 bot.

ضع هذا الملف بجانب:
- bot_v55_fix_board_stats_fahad.py
- exit_scenarios_runner.py

ثم شغله بدل V55:
python bot_v56_final_fahad.py

التعديل يعمل كطبقة آمنة فوق V55 بدون لمس الفانتزي أو ترتيب المجموعات.
"""

import os
import re
import sys
import json
import importlib.util
from datetime import datetime

BASE_FILE = os.getenv("MASIF_BOT_V55_FILE", "bot_v55_fix_board_stats_fahad.py")


def _load_v55(path: str):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"لم أجد ملف V55: {path}\n"
            "ضع bot_v55_fix_board_stats_fahad.py بجانب bot_v56_final_fahad.py "
            "أو عرّف المسار عبر MASIF_BOT_V55_FILE."
        )
    spec = importlib.util.spec_from_file_location("masif_bot_v55_base", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"تعذر تحميل ملف البوت من: {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


_AR_GROUP_TO_CODE = {
    "أ": "A", "ا": "A", "A": "A",
    "ب": "B", "B": "B",
    "ج": "C", "C": "C",
    "د": "D", "D": "D",
    "هـ": "E", "ه": "E", "E": "E",
    "و": "F", "F": "F",
    "ز": "G", "G": "G",
    "ح": "H", "H": "H",
    "ط": "I", "I": "I",
    "ي": "J", "J": "J",
    "ك": "K", "K": "K",
    "ل": "L", "L": "L",
}
_CODE_TO_AR = {"A":"أ", "B":"ب", "C":"ج", "D":"د", "E":"هـ", "F":"و", "G":"ز", "H":"ح", "I":"ط", "J":"ي", "K":"ك", "L":"ل"}
_AR_RANK = {1: "الأول", 2: "الثاني", 3: "الثالث", 4: "الرابع"}

# جدول الجولة الأخيرة — توقيت السعودية حسب جدول البطولة المعتمد داخل V55.
_V56_FINAL_ROUND = [
    ("24/06/2026", "الأربعاء", "10:00 م", "سويسرا", "كندا", "المجموعة ب"),
    ("24/06/2026", "الأربعاء", "10:00 م", "البوسنة والهرسك", "قطر", "المجموعة ب"),
    ("24/06/2026", "الأربعاء", "3:00 فجراً", "اسكتلندا", "البرازيل", "المجموعة ج"),
    ("24/06/2026", "الأربعاء", "3:00 فجراً", "المغرب", "هايتي", "المجموعة ج"),
    ("24/06/2026", "الأربعاء", "6:00 صباحاً", "التشيك", "المكسيك", "المجموعة أ"),
    ("24/06/2026", "الأربعاء", "6:00 صباحاً", "جنوب أفريقيا", "كوريا الجنوبية", "المجموعة أ"),

    ("25/06/2026", "الخميس", "10:00 م", "الإكوادور", "ألمانيا", "المجموعة هـ"),
    ("25/06/2026", "الخميس", "10:00 م", "كوراساو", "ساحل العاج", "المجموعة هـ"),
    ("25/06/2026", "الخميس", "3:00 فجراً", "تونس", "هولندا", "المجموعة و"),
    ("25/06/2026", "الخميس", "3:00 فجراً", "اليابان", "السويد", "المجموعة و"),
    ("25/06/2026", "الخميس", "6:00 صباحاً", "تركيا", "أمريكا", "المجموعة د"),
    ("25/06/2026", "الخميس", "6:00 صباحاً", "باراغواي", "أستراليا", "المجموعة د"),

    ("26/06/2026", "الجمعة", "10:00 م", "النرويج", "فرنسا", "المجموعة ط"),
    ("26/06/2026", "الجمعة", "10:00 م", "السنغال", "العراق", "المجموعة ط"),
    ("26/06/2026", "الجمعة", "3:00 فجراً", "الأوروغواي", "إسبانيا", "المجموعة ح"),
    ("26/06/2026", "الجمعة", "3:00 فجراً", "الرأس الأخضر", "السعودية", "المجموعة ح"),
    ("26/06/2026", "الجمعة", "6:00 صباحاً", "نيوزيلندا", "بلجيكا", "المجموعة ز"),
    ("26/06/2026", "الجمعة", "6:00 صباحاً", "مصر", "إيران", "المجموعة ز"),

    ("27/06/2026", "السبت", "12:00 صباحاً", "بنما", "إنجلترا", "المجموعة ل"),
    ("27/06/2026", "السبت", "12:00 صباحاً", "كرواتيا", "غانا", "المجموعة ل"),
    ("27/06/2026", "السبت", "2:30 صباحاً", "كولومبيا", "البرتغال", "المجموعة ك"),
    ("27/06/2026", "السبت", "2:30 صباحاً", "الكونغو الديمقراطية", "أوزبكستان", "المجموعة ك"),
    ("27/06/2026", "السبت", "5:00 صباحاً", "الأردن", "الأرجنتين", "المجموعة ي"),
    ("27/06/2026", "السبت", "5:00 صباحاً", "الجزائر", "النمسا", "المجموعة ي"),
]

# المسارات المعروفة من جدول دور الـ32. إن لم يتحدد الخصم، نعرض الخانة فقط.
_V56_R32_SLOTS = [
    ("2A", "2B", "28/06"),
    ("1C", "2F", "29/06"),
    ("1E", "3A/B/C/D/F", "29/06"),
    ("1F", "2C", "29/06"),
    ("2E", "2I", "30/06"),
    ("1I", "3C/D/F/G/H", "30/06"),
    ("1A", "3C/E/F/H/I", "30/06"),
    ("1L", "3E/H/I/J/K", "01/07"),
    ("1G", "3A/E/H/I/J", "01/07"),
    ("1D", "3B/E/F/I/J", "01/07"),
]

_MENU_TEXTS = {
    "📺 مباشر الآن", "🏆 لوحة البطولة", "✅ كيف تتأهل؟", "🧮 حاسبة التأهل",
    "📊 إحصائيات البطولة", "📅 المباريات القادمة", "🗂️ أرشيف البطولة", "🎮 فانتزي",
    "🥉 أفضل الثوالث", "🥉 أفضل الثوالث الآن", "✅ المتأهلون", "❌ المغادرون",
    "📋 نتائج المباريات", "🎬 ملخصات المباريات", "🏆 هدافين البطولة", "⚽ مسجلو الأهداف",
    "🏟️ مباريات دور الـ32", "🏟️ مباريات الـ32", "👀 وش أتابع الجولة الأخيرة؟",
    "⬅️ رجوع", "⬅️ الرئيسية", "🔄 تحديث الآن", "القائمة", "ابدأ", "ابدا", "المصيف",
}


def install_v56(bot):
    InlineKeyboardButton = bot.InlineKeyboardButton
    InlineKeyboardMarkup = bot.InlineKeyboardMarkup
    ReplyKeyboardMarkup = bot.ReplyKeyboardMarkup
    ContextTypes = getattr(bot, "ContextTypes", None)
    Update = getattr(bot, "Update", None)

    def _norm(x):
        try:
            return bot.normalize_name(x)
        except Exception:
            return re.sub(r"\s+", " ", str(x or "")).strip()

    def _safe_int(v, default=0):
        try:
            if isinstance(v, str):
                v = v.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
                m = re.search(r"[-+]?\d+", v)
                return int(m.group(0)) if m else default
            return int(v or default)
        except Exception:
            return default

    def _now():
        try:
            return bot._now_riyadh_text()
        except Exception:
            return datetime.now().strftime("%Y-%m-%d %H:%M")

    def _snapshot(force=False):
        try:
            return bot._v33_snapshot(force)
        except Exception:
            try:
                return bot._v32_snapshot(force)
            except Exception:
                return {}

    def _find_team(team):
        try:
            return bot._v33_find_team(team) or _norm(team)
        except Exception:
            return _norm(team)

    def _team_key(team):
        try:
            return bot._v33_team_button_key(team)
        except Exception:
            return _norm(team).replace("|", "_")

    def _team_from_key(key):
        try:
            return bot._v33_team_from_button_key(key)
        except Exception:
            return _norm(key).replace("_", " ")

    def _group_code_from_any(group):
        s = str(group or "").strip()
        if not s:
            return ""
        # المجموعة أ / Group A / A
        m = re.search(r"المجموعة\s*([أابجدهـوزحطيكل])", s, re.I)
        if m:
            return _AR_GROUP_TO_CODE.get(m.group(1), "")
        m = re.search(r"\b([A-L])\b", s, re.I)
        if m:
            return m.group(1).upper()
        for ar, code in _AR_GROUP_TO_CODE.items():
            if ar and ar in s:
                return code
        return ""

    def _group_label(code):
        code = _group_code_from_any(code) or str(code or "")
        try:
            return bot._v33_group_label(code)
        except Exception:
            return f"المجموعة {_CODE_TO_AR.get(code, code)}" if code else "المجموعة"

    def _group_rows(code, snap=None):
        snap = snap or _snapshot(False)
        try:
            teams = bot._v33_group_teams(code)
        except Exception:
            teams = []
        # الأضمن: استعمل دالة الصف لكل فريق ثم رتّب حسب المركز الحالي.
        rows = []
        for t in teams or []:
            try:
                c, pos, row, _ = bot._v33_group_row(t, snap)
                if row and (not code or _group_code_from_any(c) == _group_code_from_any(code)):
                    item = dict(row)
                    item.setdefault("team", t)
                    item["_pos"] = _safe_int(pos, 99)
                    rows.append(item)
            except Exception:
                pass
        if rows:
            return sorted(rows, key=lambda r: _safe_int(r.get("_pos"), 99))
        # بدائل من شكل snapshot إن وجد.
        groups = snap.get("groups") or snap.get("standings") or []
        for g in groups if isinstance(groups, list) else []:
            if not isinstance(g, (list, tuple, dict)):
                continue
            title = g.get("group") or g.get("title") if isinstance(g, dict) else (g[0] if g else "")
            if _group_code_from_any(title) != _group_code_from_any(code):
                continue
            rs = g.get("rows") if isinstance(g, dict) else (g[1] if len(g) > 1 else [])
            return list(rs or [])
        return []

    def _group_row(team, snap=None):
        snap = snap or _snapshot(False)
        try:
            return bot._v33_group_row(team, snap)
        except Exception:
            team_n = _find_team(team)
            for code in list(_CODE_TO_AR.keys()):
                for idx, r in enumerate(_group_rows(code, snap), start=1):
                    rt = r.get("team") or r.get("name") or r.get("Team") if isinstance(r, dict) else ""
                    if _norm(rt) == _norm(team_n):
                        return code, idx, r, _group_rows(code, snap)
        return "", None, None, []

    def _row_team(row):
        if not isinstance(row, dict):
            return ""
        return row.get("team") or row.get("name") or row.get("Team") or row.get("country") or ""

    def _status_text(team, snap=None):
        snap = snap or _snapshot(False)
        try:
            return str((snap.get("status") or {}).get(team, "") or "")
        except Exception:
            return ""

    def _is_official_qualified(status):
        try:
            return bot._v38d_is_official_qualified(status)
        except Exception:
            return ("رسمي" in str(status) and ("متأهل" in str(status) or "ضمن" in str(status)))

    def _is_official_eliminated(status):
        return "رسمي" in str(status) and ("مستبعد" in str(status) or "غادر" in str(status) or "خارج" in str(status))

    def _possible_positions(team, snap=None):
        snap = snap or _snapshot(False)
        try:
            possible, _examples = bot._v38d_possible_positions(team, snap)
            return set(int(x) for x in (possible or []) if str(x).isdigit())
        except Exception:
            code, pos, row, _rows = _group_row(team, snap)
            return {int(pos)} if pos else set()

    def _rank_word(pos):
        try:
            return bot._v38d_rank_word(pos)
        except Exception:
            return _AR_RANK.get(_safe_int(pos), f"المركز {pos}")

    def _third_is_official(team, row, snap):
        try:
            return bot._v50_third_is_official(team, row, snap)
        except Exception:
            status = " ".join([
                str((row or {}).get("third_status") or ""),
                str((row or {}).get("status") or ""),
                _status_text(team, snap),
            ])
            return ("رسمي" in status and ("أفضل" in status or "ثالث" in status) and ("ضمن" in status or "متأهل" in status))

    def _qualification_position_state(team, snap=None):
        snap = snap or _snapshot(False)
        team = _find_team(team)
        code, pos, row, rows = _group_row(team, snap)
        status = _status_text(team, snap)
        possible = _possible_positions(team, snap)
        played = _safe_int((row or {}).get("P") or (row or {}).get("played"), 0)
        pos_i = _safe_int(pos, 0)
        official = _is_official_qualified(status) or team in set(_qualified_teams(snap))
        if not official:
            return {"team": team, "code": code, "pos": pos_i, "row": row, "status": status, "official": False, "secured": False, "possible": possible}

        # أفضل ثالث رسمي له صيغة مستقلة.
        if pos_i == 3 and (_third_is_official(team, row or {}, snap) or "أفضل" in status):
            return {"team": team, "code": code, "pos": 3, "row": row, "status": status, "official": True, "secured": True, "best_third": True, "possible": {3}}

        # المركز مضمون إذا الحسبة ترجّع مركزًا واحدًا، أو انتهت المجموعة، أو نص الحالة يذكر المركز بوضوح.
        secured_by_text = any(x in status for x in ["المركز الأول", "المركز الثاني", "المركز الثالث", "كأول", "كثاني", "كثالث"])
        all_played = bool(rows) and all(_safe_int((rr or {}).get("P") or (rr or {}).get("played"), 0) >= 3 for rr in rows if isinstance(rr, dict))
        secured = (pos_i in (1, 2, 3)) and (secured_by_text or all_played or possible == {pos_i})
        return {"team": team, "code": code, "pos": pos_i, "row": row, "status": status, "official": True, "secured": secured, "best_third": False, "possible": possible}

    def _qualified_teams(snap=None):
        snap = snap or _snapshot(False)
        teams = []
        try:
            teams.extend(bot._status_effective_teams("qualified"))
        except Exception:
            pass
        try:
            teams.extend(snap.get("qualified") or [])
        except Exception:
            pass
        status = snap.get("status") if isinstance(snap, dict) else {}
        if isinstance(status, dict):
            for team, st in status.items():
                if _is_official_qualified(st):
                    teams.append(team)
        out = []
        for t in teams:
            t = _find_team(t)
            if t and t not in out:
                out.append(t)
        return out[:32]

    def _candidate_thirds(snap=None):
        snap = snap or _snapshot(False)
        return list((snap or {}).get("thirds") or [])

    def _prob_emoji(p):
        if p >= 85: return "🟢"
        if p >= 60: return "🟡"
        if p >= 35: return "🟠"
        return "🔴"

    def _third_probability(team, row, snap, rank):
        if _third_is_official(team, row, snap):
            return 100, "🟢"
        status = " ".join([str((row or {}).get("third_status") or ""), str((row or {}).get("status") or ""), _status_text(team, snap)])
        if _is_official_eliminated(status) or "مستحيل" in status:
            return 0, "🔴"
        pts = _safe_int((row or {}).get("Pts") or (row or {}).get("pts") or (row or {}).get("points"), 0)
        gd = _safe_int((row or {}).get("GD") or (row or {}).get("gd"), 0)
        gf = _safe_int((row or {}).get("GF") or (row or {}).get("gf"), 0)
        played = _safe_int((row or {}).get("P") or (row or {}).get("played"), 0)
        remaining_bonus = max(0, 3 - played) * 8
        if rank <= 8:
            p = 62 + (8 - rank) * 3 + pts * 4 + gd * 3 + min(gf, 6) + remaining_bonus
            p = max(52, min(97, p))
        else:
            p = 46 - (rank - 8) * 6 + pts * 6 + gd * 3 + min(gf, 5) + remaining_bonus
            p = max(1, min(78, p))
        return int(p), _prob_emoji(int(p))

    def _v56_best_thirds_text(force=False):
        snap = _snapshot(force)
        thirds = _candidate_thirds(snap)
        lines = ["🥉 أفضل الثوالث الآن", f"آخر تحديث: {(snap or {}).get('updated_at','-')}", ""]
        if not thirds:
            return "🥉 أفضل الثوالث الآن\nلا توجد بيانات كافية حتى الآن."
        official_count = 0
        for r in thirds[:8]:
            team = (r.get("team") or r.get("name") or "-") if isinstance(r, dict) else "-"
            if _third_is_official(team, r, snap):
                official_count += 1
        lines.append(f"📊 داخلون حاليًا ضمن أفضل 8 ثوالث: {min(8, len(thirds[:8]))}/8")
        lines.append(f"✅ ضمنوا رسميًا كأفضل ثالث: {official_count}/8")
        lines.append("")
        lines.append("✅ داخلون حاليًا:")
        for i, r in enumerate(thirds[:8], start=1):
            if not isinstance(r, dict):
                continue
            team = r.get("team") or r.get("name") or "-"
            group = r.get("group") or r.get("Group") or ""
            pts = r.get("Pts") or r.get("pts") or r.get("points") or 0
            played = r.get("P") or r.get("played") or 0
            gd = _safe_int(r.get("GD") or r.get("gd"), 0)
            gf = r.get("GF") or r.get("gf") or "-"
            official = _third_is_official(team, r, snap)
            status = "✅ ضمن أفضل ثالث رسميًا" if official else "⏳ داخل حاليًا ولم يضمن رسميًا"
            pct, em = _third_probability(team, r, snap, i)
            lines.append(f"{i}. {team}{' ✅' if official else ''} — {_group_label(group)}")
            lines.append(f"   النقاط: {pts} | لعب: {played}/3 | الفارق: {gd:+d} | الأهداف: {gf}")
            lines.append(f"   الحالة: {status}")
            lines.append(f"   فرصة التأهل كأفضل ثالث: {pct}% {em}")
            lines.append("")
        outside = thirds[8:]
        lines.append("❌ خارجون حاليًا:")
        if not outside:
            lines.append("لا يوجد.")
        for i, r in enumerate(outside, start=9):
            if not isinstance(r, dict):
                continue
            team = r.get("team") or r.get("name") or "-"
            group = r.get("group") or r.get("Group") or ""
            pts = r.get("Pts") or r.get("pts") or r.get("points") or 0
            played = r.get("P") or r.get("played") or 0
            gd = _safe_int(r.get("GD") or r.get("gd"), 0)
            gf = r.get("GF") or r.get("gf") or "-"
            official = _third_is_official(team, r, snap)
            status = "✅ ضمن أفضل ثالث رسميًا" if official else "⚠️ خارج حاليًا لكنه غير مستبعد رسميًا"
            pct, em = _third_probability(team, r, snap, i)
            lines.append(f"{i}. {team}{' ✅' if official else ''} — {_group_label(group)}")
            lines.append(f"   النقاط: {pts} | لعب: {played}/3 | الفارق: {gd:+d} | الأهداف: {gf}")
            lines.append(f"   الحالة: {status}")
            lines.append(f"   فرصة التأهل كأفضل ثالث: {pct}% {em}")
            lines.append("")
        lines.append("⚠️ أفضل 8 ثوالث من أصل 12 يتأهلون لدور الـ32.")
        lines.append("النسبة تقديرية حسب الوضع الحالي، ولا تصبح 100% إلا عند الضمان الحسابي الرسمي.")
        return "\n".join(lines).strip()

    bot._v32_best_thirds_text = _v56_best_thirds_text

    def _v56_how_qualify_text(team, force=False):
        prev = getattr(bot, "_V56_PREV_HOW_TEXT", None)
        if callable(prev):
            try:
                txt = prev(team, force)
            except Exception:
                txt = ""
        else:
            txt = ""
        if not txt:
            team2 = _find_team(team)
            if not team2:
                return "ما عرفت المنتخب. اكتب اسم منتخب مثل: السعودية أو مصر أو العراق."
            txt = f"✅ كيف يتأهل {team2}؟"
        # إذا المنتخب متأهل رسميًا، لا نحذف تفاصيله.
        snap = _snapshot(force)
        team2 = _find_team(team)
        if _is_official_qualified(_status_text(team2, snap)):
            return txt
        cut_patterns = ["اضغط الأزرار بالأسفل", "🥇 احتمالات المركز الأول", "🥈 احتمالات المركز الثاني", "🥉 احتمالات أفضل ثالث", "🏟️ الخصوم المحتملين", "📊 تفاصيل المجموعة"]
        lines = txt.splitlines()
        keep = []
        cut = False
        for line in lines:
            if any(line.strip().startswith(p) for p in cut_patterns):
                cut = True
            if not cut:
                keep.append(line)
        base = "\n".join(keep).rstrip()
        return (base + "\n\n"
                "احتمالات التأهل بـ:\n"
                "🥇 المركز الأول | 🥈 المركز الثاني | 🥉 المركز الثالث\n\n"
                "معلومات إضافية:\n"
                "🏟️ الخصوم | ❌ المغادرة | 📊 المجموعة | 🧮 سيناريو").strip()

    bot._V56_PREV_HOW_TEXT = getattr(bot, "_v33_how_qualify_text", None)
    bot._v33_how_qualify_text = _v56_how_qualify_text

    def _v56_how_qualify_keyboard(team):
        team = _find_team(team)
        snap = _snapshot(False)
        _code, pos, row, _rows = _group_row(team, snap)
        key = _team_key(team)
        status = _status_text(team, snap)
        if _is_official_qualified(status) or _is_official_eliminated(status):
            return InlineKeyboardMarkup([
                [InlineKeyboardButton("📊 المجموعة", callback_data=f"v32|how_group|{key}")],
                [InlineKeyboardButton("⬅️ رجوع", callback_data="mainmenu|home")],
            ])
        possible = _possible_positions(team, snap) or ({_safe_int(pos)} if pos else set())
        rows_btn = []
        first_row = []
        if 1 in possible:
            first_row.append(InlineKeyboardButton("🥇 المركز الأول", callback_data=f"v32|how_first|{key}"))
        else:
            first_row.append(InlineKeyboardButton("🥇 المركز الأول", callback_data=f"v32|how_first|{key}"))
        if 2 in possible:
            first_row.append(InlineKeyboardButton("🥈 المركز الثاني", callback_data=f"v32|how_second|{key}"))
        else:
            first_row.append(InlineKeyboardButton("🥈 المركز الثاني", callback_data=f"v32|how_second|{key}"))
        rows_btn.append(first_row[:2])
        rows_btn.append([
            InlineKeyboardButton("🥉 المركز الثالث", callback_data=f"v32|how_thirds|{key}"),
            InlineKeyboardButton("🏟️ الخصوم", callback_data=f"v32|opp_all|{key}"),
        ])
        rows_btn.append([
            InlineKeyboardButton("❌ المغادرة", callback_data=f"v32|how_exit|{key}"),
            InlineKeyboardButton("📊 المجموعة", callback_data=f"v32|how_group|{key}"),
        ])
        rows_btn.append([
            InlineKeyboardButton("🧮 سيناريو", callback_data="v32|calc_start"),
            InlineKeyboardButton("⬅️ رجوع", callback_data="mainmenu|home"),
        ])
        return InlineKeyboardMarkup(rows_btn)

    bot._v33_how_qualify_keyboard = _v56_how_qualify_keyboard

    def _qualified_line(info):
        team = info["team"]
        code = info.get("code") or ""
        row = info.get("row") or {}
        pts = row.get("Pts") or row.get("pts") or row.get("points") or "-" if isinstance(row, dict) else "-"
        gd = _safe_int((row or {}).get("GD") or (row or {}).get("gd"), 0) if isinstance(row, dict) else 0
        if info.get("best_third"):
            status_line = "✅ تأهل رسميًا — كأفضل ثالث"
        elif info.get("secured") and info.get("pos") in (1, 2, 3):
            status_line = f"✅ تأهل رسميًا — المركز {_AR_RANK.get(info.get('pos'), info.get('pos'))}"
        else:
            poss = sorted([p for p in info.get("possible", set()) if p in (1,2,3)])
            poss_txt = " أو ".join(_AR_RANK.get(p, str(p)) for p in poss) if poss else "يتحدد لاحقًا"
            status_line = f"✅ تأهل رسميًا — المركز لم يُحسم بعد\n   المركز المحتمل: {poss_txt}"
        extra = f"   المجموعة: {_group_label(code)}"
        if pts != "-":
            extra += f" | النقاط: {pts} | الفارق: {gd:+d}"
        return f"{team}\n   {status_line}\n{extra}"

    def _v56_qualified_text(force=False):
        snap = _snapshot(force)
        teams = _qualified_teams(snap)
        lines = ["✅ المتأهلون رسميًا إلى دور الـ32", f"آخر تحديث: {(snap or {}).get('updated_at','-')}", f"العدد: {len(teams)}/32", ""]
        if not teams:
            lines.append("لا توجد منتخبات متأهلة رسميًا حتى الآن.")
            return "\n".join(lines).strip()
        groups = {"first": [], "second": [], "third": [], "pending": []}
        for t in teams:
            info = _qualification_position_state(t, snap)
            if info.get("best_third") or (info.get("secured") and info.get("pos") == 3):
                groups["third"].append(info)
            elif info.get("secured") and info.get("pos") == 1:
                groups["first"].append(info)
            elif info.get("secured") and info.get("pos") == 2:
                groups["second"].append(info)
            else:
                groups["pending"].append(info)
        sections = [
            ("🥇 ضمنوا المركز الأول رسميًا:", groups["first"]),
            ("🥈 ضمنوا المركز الثاني رسميًا:", groups["second"]),
            ("🥉 ضمنوا المركز الثالث/أفضل ثالث رسميًا:", groups["third"]),
            ("⏳ تأهلوا رسميًا لكن المركز لم يُحسم بعد:", groups["pending"]),
        ]
        any_section = False
        for title, arr in sections:
            if not arr:
                continue
            any_section = True
            lines.append(title)
            for i, info in enumerate(arr, 1):
                block = _qualified_line(info).splitlines()
                lines.append(f"{i}. {block[0]}")
                lines.extend(block[1:])
                lines.append("")
        if not any_section:
            for i, t in enumerate(teams, 1):
                lines.append(f"{i}. {t}\n   ✅ تأهل رسميًا — المركز لم يُحسم بعد")
        return "\n".join(lines).strip()

    def _v56_qualified_kb(back_to="mainmenu|home"):
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("🏟️ مباريات دور الـ32", callback_data="v32|round32")],
            [InlineKeyboardButton("🔄 تحديث الآن", callback_data="v32|qualified_force"), InlineKeyboardButton("⬅️ رجوع", callback_data=back_to)],
        ])

    async def _v56_send_qualified(message, force=False):
        await message.reply_text(_v56_qualified_text(force), reply_markup=_v56_qualified_kb())

    bot._v56_qualified_text = _v56_qualified_text

    def _slot_label(slot):
        slot = str(slot or "").strip().upper()
        m = re.fullmatch(r"([123])([A-L])", slot)
        if m:
            rank = {"1": "أول", "2": "ثاني", "3": "ثالث"}.get(m.group(1), m.group(1))
            return f"{rank} المجموعة {_CODE_TO_AR.get(m.group(2), m.group(2))}"
        m = re.fullmatch(r"3([A-L](?:/[A-L])*)", slot)
        if m:
            groups = "/".join(_CODE_TO_AR.get(x, x) for x in m.group(1).split("/"))
            return f"ثالث المجموعة {groups}"
        return slot

    def _slot_team(slot, snap=None):
        snap = snap or _snapshot(False)
        slot = str(slot or "").strip().upper()
        m = re.fullmatch(r"([12])([A-L])", slot)
        if m:
            rank = int(m.group(1)); code = m.group(2)
            rows = _group_rows(code, snap)
            if len(rows) >= rank:
                team = _row_team(rows[rank-1])
                if team:
                    info = _qualification_position_state(team, snap)
                    if info.get("secured") and info.get("pos") == rank:
                        return team, f"{team} ضمن المركز {_AR_RANK.get(rank)} رسميًا", True
            return _slot_label(slot), "الخانة لم تكتمل رسميًا", False
        m = re.fullmatch(r"3([A-L](?:/[A-L])*)", slot)
        if m:
            allowed = set(m.group(1).split("/"))
            for r in _candidate_thirds(snap):
                if not isinstance(r, dict):
                    continue
                code = _group_code_from_any(r.get("group") or r.get("Group") or "")
                if code in allowed:
                    return _slot_label(slot), f"الخصم المتوقع: {r.get('team') or r.get('name') or '-'}", False
            return _slot_label(slot), "ينتظر ترتيب أفضل الثوالث", False
        return _slot_label(slot), "", False

    def _v56_expected_third(slot, snap=None):
        snap = snap or _snapshot(False)
        slot = str(slot or "").upper()
        m = re.fullmatch(r"3([A-L](?:/[A-L])*)", slot)
        if not m:
            return ""
        allowed = set(m.group(1).split("/"))
        for r in _candidate_thirds(snap):
            if not isinstance(r, dict):
                continue
            code = _group_code_from_any(r.get("group") or r.get("Group") or "")
            team = r.get("team") or r.get("name") or ""
            if code in allowed and team:
                return team
        return ""

    def _v56_round32_text(force=False):
        snap = _snapshot(force)
        lines = ["🏟️ مباريات دور الـ32", f"آخر تحديث: {(snap or {}).get('updated_at','-')}", ""]
        lines.append("✅ المسارات المعروفة حتى الآن:")
        lines.append("")
        for i, (home_slot, away_slot, date_s) in enumerate(_V56_R32_SLOTS, 1):
            home, h_status, home_ok = _slot_team(home_slot, snap)
            away, a_status, away_ok = _slot_team(away_slot, snap)
            lines.append(f"{i}. {home} × {away}")
            if date_s:
                lines.append(f"   الموعد: {date_s}")
            if home_ok or away_ok:
                status_parts = []
                if home_ok: status_parts.append(h_status)
                if away_ok: status_parts.append(a_status)
                lines.append("   الحالة: " + " | ".join(status_parts))
            if str(home_slot).startswith("3") or str(away_slot).startswith("3"):
                exp = _v56_expected_third(home_slot, snap) or _v56_expected_third(away_slot, snap)
                if exp:
                    lines.append(f"   🔎 حسب ترتيب أفضل الثوالث الآن: الخصم المتوقع {exp}")
                    lines.append("   ⚠️ غير مؤكد حتى الآن — يتغير حسب نتائج الجولة الأخيرة وتركيبة أفضل الثوالث")
                else:
                    lines.append("   ⚠️ خصم أفضل الثوالث غير مؤكد حتى الآن")
            elif home_ok and away_ok:
                lines.append("   الحالة: مباراة مؤكدة")
            else:
                if not (home_ok or away_ok):
                    lines.append("   الحالة: المسار معروف والخانات تنتظر حسم المراكز")
            lines.append("")
        lines.append("ملاحظة: إذا لم يتحدد الخصم بالاسم، تظهر خانته فقط إلى أن تُحسم نتائج الجولة الأخيرة.")
        return "\n".join(lines).strip()

    def _match_importance(t1, t2, group, snap=None):
        snap = snap or _snapshot(False)
        code = _group_code_from_any(group)
        rows = _group_rows(code, snap)
        row_map = {_norm(_row_team(r)): r for r in rows if isinstance(r, dict)}
        r1 = row_map.get(_norm(_find_team(t1))) or row_map.get(_norm(t1)) or {}
        r2 = row_map.get(_norm(_find_team(t2))) or row_map.get(_norm(t2)) or {}
        pts1, pts2 = _safe_int(r1.get("Pts") or r1.get("pts"), 0), _safe_int(r2.get("Pts") or r2.get("pts"), 0)
        pos1 = next((i for i, r in enumerate(rows, 1) if _norm(_row_team(r)) in {_norm(t1), _norm(_find_team(t1))}), 0)
        pos2 = next((i for i, r in enumerate(rows, 1) if _norm(_row_team(r)) in {_norm(t2), _norm(_find_team(t2))}), 0)
        max_pts = max(pts1, pts2)
        min_pts = min(pts1, pts2)
        close = (max_pts - min_pts) <= 3
        score = 30
        reasons = []
        if pos1 in (1,2,3) or pos2 in (1,2,3):
            score += 25; reasons.append("تؤثر على مراكز التأهل المباشر أو المركز الثالث")
        if close:
            score += 18; reasons.append("الفوارق النقطية قريبة")
        if any(p in (3,4) for p in [pos1, pos2]):
            score += 15; reasons.append("مرتبطة بسباق أفضل الثوالث/المغادرة")
        if max_pts >= 4:
            score += 8; reasons.append("فيها منتخب قريب من الحسم")
        if not reasons:
            reasons.append("تأثيرها أقل على شكل المجموعة الحالي")
        if score >= 78:
            level = "🔥 عالية جدًا"
        elif score >= 60:
            level = "🟠 مؤثرة"
        elif score >= 45:
            level = "🟡 متابعة"
        else:
            level = "⚪ قليلة التأثير"
        return score, level, "، ".join(reasons[:2])

    def _v56_last_round_watch_text(force=False):
        snap = _snapshot(force)
        by_day = {}
        for d, day, time_s, t1, t2, group in _V56_FINAL_ROUND:
            by_day.setdefault((d, day), {}).setdefault(time_s, []).append((t1, t2, group))
        lines = ["👀 وش أتابع الجولة الأخيرة؟", "من الأربعاء إلى السبت — الأحد يبدأ دور الـ32", f"آخر تحديث: {(snap or {}).get('updated_at','-')}", ""]
        day_no = 0
        for (d, day), times in sorted(by_day.items(), key=lambda x: datetime.strptime(x[0][0], "%d/%m/%Y")):
            day_no += 1
            lines.append(f"📅 {day} {d} — اليوم {day_no} من الجولة الأخيرة")
            for time_s, matches in times.items():
                lines.append(f"⏰ {time_s}")
                scored = []
                for t1, t2, group in matches:
                    score, level, reason = _match_importance(t1, t2, group, snap)
                    scored.append((score, t1, t2, group, level, reason))
                    lines.append(f"- {t1} × {t2} — {group}")
                    lines.append(f"  الأهمية: {level}")
                    lines.append(f"  السبب: {reason}")
                if scored:
                    best = sorted(scored, reverse=True, key=lambda x: x[0])[0]
                    lines.append(f"🎯 اقتراح المصيف: تابع {best[1]} × {best[2]}")
                    lines.append(f"   لأنها الأكثر تأثيرًا في هذا التوقيت حسب وضع المجموعة وأفضل الثوالث.")
                lines.append("")
        lines.append("التقييم يتغير تلقائيًا بعد كل نتيجة حسب النقاط، الفارق، المركز، والمنافسة على أفضل الثوالث.")
        return "\n".join(lines).strip()

    def _v56_how_exit_text(team):
        team = _find_team(team)
        snap = _snapshot(False)
        code, pos, row, rows = _group_row(team, snap)
        status = _status_text(team, snap) or "ما زال ينافس"
        lines = [f"❌ احتمال مغادرة {team}", "", f"الحالة الحالية: {status}"]
        if row:
            lines.append(f"المركز الحالي: {pos} | النقاط: {row.get('Pts')} | الفارق: {_safe_int(row.get('GD')):+d}")
        if _is_official_eliminated(status):
            lines.append("\n❌ المنتخب مستبعد/مغادر رسميًا حسب الحسبة الحالية.")
        else:
            lines.append("\nلم يُحسم خروجه رسميًا إلا إذا انتهت فرصه حسابيًا من المركزين الأول والثاني ومن أفضل الثوالث.")
        return "\n".join(lines).strip()

    # ---------- القوائم الرئيسية ----------
    def _public_main_reply_keyboard():
        return ReplyKeyboardMarkup(
            [
                ["📺 مباشر الآن", "🏆 لوحة البطولة"],
                ["✅ كيف تتأهل؟", "🧮 حاسبة التأهل"],
                ["📊 إحصائيات البطولة", "📅 المباريات القادمة"],
                ["🏟️ مباريات دور الـ32", "🗂️ أرشيف البطولة"],
                ["🎮 فانتزي"],
                ["👀 وش أتابع الجولة الأخيرة؟"],
            ],
            resize_keyboard=True,
            one_time_keyboard=False,
            input_field_placeholder="اكتب اسم منتخب أو اختر من القائمة",
        )

    def _public_main_keyboard():
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("📺 مباشر الآن", callback_data="mainmenu|live"), InlineKeyboardButton("🏆 لوحة البطولة", callback_data="v32|board")],
            [InlineKeyboardButton("✅ كيف تتأهل؟", callback_data="v32|how_start"), InlineKeyboardButton("🧮 حاسبة التأهل", callback_data="v32|calc_start")],
            [InlineKeyboardButton("📊 إحصائيات البطولة", callback_data="v32|stats"), InlineKeyboardButton("📅 المباريات القادمة", callback_data="mainmenu|fixtures")],
            [InlineKeyboardButton("🏟️ مباريات دور الـ32", callback_data="mainmenu|round32"), InlineKeyboardButton("🗂️ أرشيف البطولة", callback_data="v32|archive")],
            [InlineKeyboardButton("🎮 فانتزي", callback_data="v32|fantasy_gate")],
            [InlineKeyboardButton("👀 وش أتابع الجولة الأخيرة؟", callback_data="mainmenu|last_round_watch")],
        ])

    bot._public_main_reply_keyboard = _public_main_reply_keyboard
    bot._public_main_keyboard = _public_main_keyboard
    try:
        bot.V32_FINAL_MENU_LABELS.update(_MENU_TEXTS)
    except Exception:
        bot.V32_FINAL_MENU_LABELS = set(_MENU_TEXTS)

    def _v32_board_keyboard():
        rows = [
            [InlineKeyboardButton("🏁 سباق التأهل", callback_data="v32|race")],
            [InlineKeyboardButton("✅ المتأهلون", callback_data="v32|qualified"), InlineKeyboardButton("❌ المستبعدين", callback_data="v32|eliminated")],
            [InlineKeyboardButton("🥉 أفضل الثوالث الآن", callback_data="v32|thirds"), InlineKeyboardButton("🔥 مباريات الحسم", callback_data="v32|decisive")],
            [InlineKeyboardButton("📊 إحصائيات البطولة", callback_data="v32|stats")],
            [InlineKeyboardButton("🔄 تحديث الآن", callback_data="v32|board_force"), InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")],
        ]
        return InlineKeyboardMarkup(rows)

    bot._v32_board_keyboard = _v32_board_keyboard

    def _v34_stats_keyboard():
        return InlineKeyboardMarkup([
            [InlineKeyboardButton("📊 ترتيب المجموعات", callback_data="mainmenu|groups"), InlineKeyboardButton("🏆 هدافين البطولة", callback_data="mainmenu|scorers")],
            [InlineKeyboardButton("🥉 أفضل الثوالث الآن", callback_data="v32|thirds"), InlineKeyboardButton("✅ المتأهلون", callback_data="v32|qualified")],
            [InlineKeyboardButton("✅❌ التأهل والمغادرة", callback_data="v32|status_home"), InlineKeyboardButton("🏟️ مباريات دور الـ32", callback_data="v32|round32")],
            [InlineKeyboardButton("⬅️ رجوع", callback_data="mainmenu|home")],
        ])

    bot._v34_stats_keyboard = _v34_stats_keyboard

    # ---------- وضع تحديث الهدافين اليدوي ----------
    def _clear_scorers_wait(context):
        try:
            for k in ["v56_waiting_manual_top_scorers", "v53_waiting_manual_top_scorers", "manual_top_scorers_waiting", "pending_manual_top_scorers", "awaiting_manual_top_scorers"]:
                context.user_data.pop(k, None)
        except Exception:
            pass

    def _set_scorers_wait(context):
        try:
            context.user_data["v56_waiting_manual_top_scorers"] = True
        except Exception:
            pass

    def _is_scorers_wait(context):
        try:
            return bool(context.user_data.get("v56_waiting_manual_top_scorers"))
        except Exception:
            return False

    def _is_cancel_text(text):
        t = _norm(text).replace("إ", "ا").replace("أ", "ا")
        return t in {"/الغاء", "الغاء", "الغاء التحديث", "إلغاء", "الغاء تحديث الهدافين"}

    def _is_menu_text(text):
        t = _norm(text)
        plain = re.sub(r"^[\W_]+", "", t).strip()
        if t in _MENU_TEXTS or plain in _MENU_TEXTS:
            return True
        if len(t) <= 40 and any(w in t for w in ["القائمة", "بطولة", "التأهل", "المباريات", "الثوالث", "هدافين", "فانتزي", "رجوع"]):
            return "|" not in t and " - " not in t
        return False

    async def v53_manual_top_scorers_command(update, context):
        msg = update.effective_message
        text = msg.text or ""
        payload = re.sub(r"^/(?:تحديث_هدافين_يدوي|تحديث_الهدافين_يدوي|تعديل_الهدافين|تعديل_هدافين)(?:@\w+)?\s*", "", text, flags=re.I).strip()
        if not payload:
            _set_scorers_wait(context)
            await msg.reply_text(
                "✏️ أرسل الهدافين برسالة واحدة فقط بهذه الصيغة:\n\n"
                "ميسي | الأرجنتين | 5\n"
                "مبابي | فرنسا | 4\n"
                "هالاند | النرويج | 4\n\n"
                "للإلغاء اكتب: /الغاء\n"
                "وأي زر من القائمة يلغي وضع الانتظار."
            )
            return
        _clear_scorers_wait(context)
        items, bad = bot._v53_parse_manual_top_scorers_text(text)
        if not items:
            await msg.reply_text("❌ ما قدرت أقرأ أي لاعب. أرسلها مثل:\nميسي | الأرجنتين | 5")
            return
        data = bot._v53_save_manual_top_scorers(items, getattr(update.effective_user, "id", None))
        preview = "\n".join([f"{i+1}. {it['name']} — {it['team']} — {it['goals']}" for i, it in enumerate(items[:12])])
        extra = f"\n\n⚠️ أسطر لم تُقرأ: {len(bad)}" if bad else ""
        await msg.reply_text(
            f"✅ تم حفظ تحديث الهدافين اليدوي.\n"
            f"عدد اللاعبين: {len(items)}\n"
            f"آخر تحديث: {data.get('updated_at')}\n\n"
            f"{preview}{extra}\n\n"
            "لعرضها اضغط: 🏆 هدافين البطولة\n"
            "وللرجوع للمصدر الرسمي استخدم: /مسح_تحديث_الهدافين"
        )

    async def v53_manual_top_scorers_text_router(update, context):
        # لا يعمل إلا إذا وضع الانتظار مفعّل، ويقبل رسالة واحدة فقط.
        if not _is_scorers_wait(context):
            return False
        try:
            if not bot.is_admin_user(update):
                _clear_scorers_wait(context)
                return False
        except Exception:
            pass
        msg = update.effective_message
        text = (msg.text if msg else "") or ""
        if _is_cancel_text(text):
            _clear_scorers_wait(context)
            await msg.reply_text("✅ تم إلغاء تحديث الهدافين اليدوي.")
            return True
        if text.startswith("/") and not re.match(r"^/(?:تحديث_هدافين_يدوي|تحديث_الهدافين_يدوي|تعديل_الهدافين|تعديل_هدافين)", text):
            _clear_scorers_wait(context)
            await msg.reply_text("✅ تم إلغاء وضع تحديث الهدافين اليدوي بسبب أمر جديد.")
            return False
        if _is_menu_text(text):
            _clear_scorers_wait(context)
            return False
        _clear_scorers_wait(context)
        items, bad = bot._v53_parse_manual_top_scorers_text(text)
        if not items:
            await msg.reply_text("❌ ما قدرت أقرأ الهدافين. أعد تشغيل الوضع بالأمر /تحديث_هدافين_يدوي ثم أرسل القائمة برسالة واحدة.")
            return True
        data = bot._v53_save_manual_top_scorers(items, getattr(update.effective_user, "id", None))
        preview = "\n".join([f"{i+1}. {it['name']} — {it['team']} — {it['goals']}" for i, it in enumerate(items[:12])])
        extra = f"\n\n⚠️ أسطر لم تُقرأ: {len(bad)}" if bad else ""
        await msg.reply_text(f"✅ تم حفظ تحديث الهدافين اليدوي.\nعدد اللاعبين: {len(items)}\nآخر تحديث: {data.get('updated_at')}\n\n{preview}{extra}")
        return True

    bot.v53_manual_top_scorers_command = v53_manual_top_scorers_command
    bot.v53_manual_top_scorers_text_router = v53_manual_top_scorers_text_router
    bot._v54_clear_manual_top_scorers_waiting = _clear_scorers_wait
    bot._v54_manual_top_scorers_waiting = _is_scorers_wait

    # ---------- callbacks / routers ----------
    _prev_v32_callback = getattr(bot, "v32_callback", None)
    _prev_public_menu_callback = getattr(bot, "public_menu_callback", None)
    _prev_public_reply_menu_router = getattr(bot, "public_reply_menu_router", None)
    _prev_qualified_list_command = getattr(bot, "qualified_list_command", None)

    async def v32_callback(update, context):
        q = update.callback_query
        if q:
            _clear_scorers_wait(context)
        data = (q.data or "") if q else ""
        parts = data.split("|")
        action = parts[1] if len(parts) > 1 else ""
        if action in {"round32", "r32"}:
            try:
                await q.answer()
            except Exception:
                pass
            await q.edit_message_text(_v56_round32_text(False), reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")]]))
            return
        if action == "last_round_watch":
            try:
                await q.answer()
            except Exception:
                pass
            await q.edit_message_text(_v56_last_round_watch_text(False), reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")]]))
            return
        if action == "qualified":
            try:
                await q.answer()
            except Exception:
                pass
            await q.edit_message_text(_v56_qualified_text(False), reply_markup=_v56_qualified_kb("v32|board"))
            return
        if action == "qualified_force":
            try:
                await q.answer()
            except Exception:
                pass
            await q.edit_message_text(_v56_qualified_text(True), reply_markup=_v56_qualified_kb("v32|board"))
            return
        if action == "how_exit" and len(parts) >= 3:
            try:
                await q.answer()
            except Exception:
                pass
            team = _team_from_key(parts[2])
            await q.edit_message_text(_v56_how_exit_text(team), reply_markup=_v56_how_qualify_keyboard(team))
            return
        if callable(_prev_v32_callback):
            return await _prev_v32_callback(update, context)

    bot.v32_callback = v32_callback

    async def public_menu_callback(update, context):
        q = update.callback_query
        if q:
            _clear_scorers_wait(context)
        data = (q.data or "") if q else ""
        if data == "mainmenu|round32":
            try:
                await q.answer()
            except Exception:
                pass
            await q.edit_message_text(_v56_round32_text(False), reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")]]))
            return
        if data == "mainmenu|last_round_watch":
            try:
                await q.answer()
            except Exception:
                pass
            await q.edit_message_text(_v56_last_round_watch_text(False), reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")]]))
            return
        if callable(_prev_public_menu_callback):
            return await _prev_public_menu_callback(update, context)

    bot.public_menu_callback = public_menu_callback

    async def public_reply_menu_router(update, context):
        msg = update.message or update.effective_message
        text = _norm(msg.text if msg else "")
        if _is_cancel_text(text):
            _clear_scorers_wait(context)
            await msg.reply_text("✅ تم الإلغاء.", reply_markup=_public_main_reply_keyboard())
            return
        if text in _MENU_TEXTS:
            _clear_scorers_wait(context)
        if text in {"🏟️ مباريات دور الـ32", "🏟️ مباريات الـ32"}:
            await msg.reply_text(_v56_round32_text(False), reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")]]))
            return
        if text == "👀 وش أتابع الجولة الأخيرة؟":
            await msg.reply_text(_v56_last_round_watch_text(False), reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")]]))
            return
        if callable(_prev_public_reply_menu_router):
            return await _prev_public_reply_menu_router(update, context)

    bot.public_reply_menu_router = public_reply_menu_router

    async def qualified_list_command(update, context):
        await update.effective_message.reply_text(_v56_qualified_text(False), reply_markup=_v56_qualified_kb("mainmenu|home"))

    bot.qualified_list_command = qualified_list_command
    # بعض النسخ تستعمل اسمًا آخر.
    bot.qualified_show_list_command = qualified_list_command

    # أوامر نصية للميزات الجديدة، لو احتجتها مباشرة.
    async def v56_round32_command(update, context):
        await update.effective_message.reply_text(_v56_round32_text(False), reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")]]))

    async def v56_last_round_watch_command(update, context):
        await update.effective_message.reply_text(_v56_last_round_watch_text(False), reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ الرئيسية", callback_data="mainmenu|home")]]))

    bot.v56_round32_command = v56_round32_command
    bot.v56_last_round_watch_command = v56_last_round_watch_command

    # Patch add_handlers to register the new command aliases if main uses add_handlers.
    _prev_add_handlers = getattr(bot, "add_handlers", None)
    if callable(_prev_add_handlers):
        def add_handlers(app):
            _prev_add_handlers(app)
            try:
                app.add_handler(bot.MessageHandler(bot.filters.TEXT & bot.filters.Regex(r"^/(?:مباريات_دور_32|مباريات_ال32|دور_32|دور32)(?:\s|$)"), v56_round32_command))
                app.add_handler(bot.MessageHandler(bot.filters.TEXT & bot.filters.Regex(r"^/(?:وش_اتابع|وش_أتابع|الجولة_الأخيرة|الجولة_الاخيرة)(?:\s|$)"), v56_last_round_watch_command))
            except Exception:
                pass
        bot.add_handlers = add_handlers

    return bot


BOT = _load_v55(BASE_FILE)
install_v56(BOT)

if __name__ == "__main__":
    if not hasattr(BOT, "main"):
        raise RuntimeError("ملف V55 لا يحتوي على main()")
    BOT.main()
