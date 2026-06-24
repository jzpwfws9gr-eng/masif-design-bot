import os, json, time, asyncio, traceback, re
from datetime import datetime, timedelta

# خدمة التحديث المنفصلة V60.4
# اعتماد فهد:
# - رجعنا طريقة السحب الأولى من ملف "ممتتتاز .py" كما هي عبر دوال bot.py الأصلية.
# - لا يوجد Timeout خارجي يوقف المرحلة بعد دقيقتين؛ كل دالة تأخذ وقتها مثل القديم.
# - يبقى timeout الداخلي لطلبات ESPN كما في الكود القديم فقط، عشان الطلب الواحد ما يعلق للأبد.
# - بوت المستخدمين لا يسحب من ESPN وقت الضغط؛ يقرأ الكاش فقط.
os.environ.setdefault('MASIF_USER_BOT_CACHE_ONLY', '0')
import bot

CACHE_FILE = os.environ.get('MASIF_SHARED_CACHE_FILE', 'masif_shared_cache.json')
INTERVAL = int(os.environ.get('MASIF_UPDATE_INTERVAL_SECONDS', '120'))
LOCK_FILE = os.environ.get('MASIF_UPDATE_LOCK_FILE', 'masif_update.lock')
FIRST_STAGE_TIMEOUT = int(os.environ.get('MASIF_FIRST_STAGE_TIMEOUT_SECONDS', '0'))  # V60.4 لا يستخدم كمهلة خارجية
STAGE_TIMEOUT = int(os.environ.get('MASIF_STAGE_TIMEOUT_SECONDS', '0'))  # V60.4 لا يستخدم كمهلة خارجية


def now_makkah(fmt='%Y-%m-%d %H:%M:%S'):
    return (datetime.utcnow() + timedelta(hours=3)).strftime(fmt)


def next_time(seconds=None):
    return (datetime.utcnow() + timedelta(hours=3, seconds=int(seconds or INTERVAL))).strftime('%H:%M:%S')


def load_cache():
    try:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
    except Exception:
        pass
    return {'blocks': {}, 'status': {}, 'raw': {}}


def save_cache(data):
    data = data or {}
    data.setdefault('blocks', {})
    data.setdefault('status', {})
    data.setdefault('raw', {})
    tmp = CACHE_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CACHE_FILE)


def set_status(**kw):
    data = load_cache()
    st = data.setdefault('status', {})
    st.update(kw)
    save_cache(data)


def put_block(key, text, meta=None):
    data = load_cache()
    blocks = data.setdefault('blocks', {})
    blocks[key] = {'text': str(text or '').strip(), 'updated_at': now_makkah(), 'meta': meta or {}}
    save_cache(data)


def put_raw(key, value):
    data = load_cache()
    raw = data.setdefault('raw', {})
    raw[key] = {'value': value, 'updated_at': now_makkah()}
    save_cache(data)


def existing_text(key):
    try:
        return (((load_cache().get('blocks') or {}).get(key) or {}).get('text') or '').strip()
    except Exception:
        return ''


async def run_sync_with_timeout(func, timeout=None):
    # V60.4: لا نستخدم wait_for هنا، لأن المستخدم طلب إلغاء مهلة الدقائق التي توقف التحديث.
    # بما أن updater.py يعمل كعملية منفصلة عن bot.py، لو طالت مرحلة التحديث لا تعلق بوت المستخدمين.
    # كل طلب ESPN داخليًا عنده timeout من دوال النسخة القديمة نفسها.
    return await asyncio.to_thread(func)


def _safe_int(x, default=0):
    try: return int(float(str(x).strip()))
    except Exception: return default


def _norm_date(d):
    try:
        return bot._normalize_date_arg(d)
    except Exception:
        return str(d or '').strip()


def _active_live_date():
    for name in ('_v41_active_live_date', '_v40_active_live_date', '_v33_active_live_date', '_v29_active_fixture_date'):
        fn = getattr(bot, name, None)
        if callable(fn):
            try:
                v = fn()
                if v: return _norm_date(v)
            except Exception:
                pass
    return now_makkah('%d/%m/%Y')


def _fixture_title(d):
    try:
        return bot._v26_fixture_title(d)
    except Exception:
        return str(d or '')


def _fixtures_for_date(d):
    d = _norm_date(d)
    for name in ('_v40_fixtures_for_date', '_fixtures_for_date'):
        fn = getattr(bot, name, None)
        if callable(fn):
            try:
                rows = fn(d) or []
                if callable(getattr(bot, '_v26_dedupe_fixture_matches', None)):
                    rows = bot._v26_dedupe_fixture_matches(rows)
                return list(rows or [])
            except Exception:
                pass
    return []


def _fixture_dates_upto_active(limit=18):
    dates = []
    try:
        dates = [d for d, _ in bot._fixture_dates()]
    except Exception:
        dates = []
    dates = [_norm_date(d) for d in dates if d]
    # خذ آخر مجموعة تواريخ بدل محاولة كل البطولة في كل دورة
    return dates[-int(limit):] if dates else [_active_live_date()]


def _espn_events_for_day(d, force=False, debug=None):
    debug = debug if debug is not None else []
    d = _norm_date(d)
    # استخدم دالة النسخة القديمة مع الكاش أولًا. لا نعمل force إلا إذا ما فيه شيء.
    try:
        evs = bot._v41_espn_day_events(d, force=False, debug=debug) if callable(getattr(bot, '_v41_espn_day_events', None)) else []
    except Exception as e:
        debug.append(f'ESPN cache day error: {type(e).__name__}')
        evs = []
    if (not evs) and force:
        try:
            evs = bot._v41_espn_day_events(d, force=True, debug=debug) if callable(getattr(bot, '_v41_espn_day_events', None)) else []
        except Exception as e:
            debug.append(f'ESPN force day error: {type(e).__name__}')
            evs = []
    return evs or []


def _team_match(a, b):
    try:
        return bool(bot._v41_team_match(a, b))
    except Exception:
        aa = str(a or '').lower().strip(); bb = str(b or '').lower().strip()
        return aa and bb and (aa in bb or bb in aa)


def _competitors_from_event(ev):
    try:
        return bot._v41_competitors_from_event(ev)
    except Exception:
        comps = ((ev.get('competitions') or [{}])[0].get('competitors') or []) if isinstance(ev, dict) else []
        out = []
        for c in comps:
            team = c.get('team') or {}
            out.append({'name': team.get('displayName') or team.get('name') or team.get('shortDisplayName') or team.get('abbreviation') or '', 'abbr': team.get('abbreviation') or '', 'score': c.get('score'), 'raw': c})
        return out


def _status_from_event(ev):
    try:
        return bot._v41_status_from_espn(ev)
    except Exception:
        st = ((ev.get('status') or {}).get('type') or {}) if isinstance(ev, dict) else {}
        detail = st.get('detail') or st.get('shortDetail') or st.get('description') or ''
        state = st.get('state') or ''
        if st.get('completed') or state == 'post': return 'FT — انتهت المباراة'
        if state == 'in': return detail or 'مباشر'
        return detail or 'لم تبدأ'


def _is_live_event(ev):
    try:
        st = ((ev.get('status') or {}).get('type') or {}) if isinstance(ev, dict) else {}
        return str(st.get('state') or '').lower() == 'in'
    except Exception:
        return False


def _is_finished_obj(obj):
    try:
        if callable(getattr(bot, '_is_finished_obj', None)):
            return bool(bot._is_finished_obj(obj))
    except Exception:
        pass
    s = str((obj or {}).get('status') or '').lower()
    return 'ft' in s or 'انتهت' in s or 'final' in s


def _numeric_score(obj):
    try:
        return bool(bot._patch6_numeric_score(obj))
    except Exception:
        try:
            int(str(obj.get('score1')).strip()); int(str(obj.get('score2')).strip()); return True
        except Exception:
            return False


def _event_matches_fixture(ev, m):
    comps = _competitors_from_event(ev)
    if len(comps) < 2:
        return None
    t1, t2 = m.get('team1'), m.get('team2')
    first = second = None
    for c in comps:
        names = [c.get('name'), c.get('abbr')]
        if any(_team_match(t1, n) for n in names): first = c
        if any(_team_match(t2, n) for n in names): second = c
    if not first or not second or first is second:
        return None
    obj = {
        'team1': t1, 'team2': t2,
        'score1': str(first.get('score') if first.get('score') is not None else '').strip(),
        'score2': str(second.get('score') if second.get('score') is not None else '').strip(),
        'status': _status_from_event(ev),
        'source': 'ESPN', 'actual_source': 'ESPN',
        'event_id': str(ev.get('id') or ''),
        'summary_url': f"https://www.espn.com/soccer/match/_/gameId/{ev.get('id')}" if ev.get('id') else '',
        'scorers': [],
        'live': _is_live_event(ev),
        'saved_at': now_makkah(),
    }
    # خزّن النتيجة في كاش النسخة القديمة إن وجدت، بدون جلب summary حتى يبقى سريع.
    try:
        if _numeric_score(obj) and callable(getattr(bot, '_v33_put_cached_match_result', None)):
            bot._v33_put_cached_match_result(m, obj, 'ESPN')
    except Exception:
        pass
    try:
        if callable(getattr(bot, '_save_fast_live_cache', None)):
            bot._save_fast_live_cache(t1, t2, _norm_date(m.get('date')), obj)
    except Exception:
        pass
    return obj


def _label_for_match(m, obj):
    t1 = bot.canonical_team_name(m.get('team1')) if callable(getattr(bot, 'canonical_team_name', None)) else m.get('team1')
    t2 = bot.canonical_team_name(m.get('team2')) if callable(getattr(bot, 'canonical_team_name', None)) else m.get('team2')
    t1 = t1 or m.get('team1') or ''
    t2 = t2 or m.get('team2') or ''
    if isinstance(obj, dict) and _numeric_score(obj):
        icon = '✅' if _is_finished_obj(obj) else ('🔴' if obj.get('live') else '⏳')
        return f"{icon} {t1} {obj.get('score1')} - {obj.get('score2')} {t2}"
    return f"⏳ {t1} × {t2}"


def _goal_lines(obj):
    if not isinstance(obj, dict) or not _numeric_score(obj):
        return []
    try:
        s1, s2 = int(obj.get('score1')), int(obj.get('score2'))
    except Exception:
        return []
    total = s1 + s2
    if total <= 0:
        return ['⚽ لا يوجد أهداف']
    scorers = [str(x).strip() for x in (obj.get('scorers') or []) if str(x).strip()]
    if not scorers:
        return ['⚽ الهدافون قيد التحديث']
    lines = ['⚽ الهدافون:']
    lines.extend(f'- {x}' for x in scorers[:12])
    return lines


def _detail_for_match(m, obj):
    t1 = bot.canonical_team_name(m.get('team1')) if callable(getattr(bot, 'canonical_team_name', None)) else m.get('team1')
    t2 = bot.canonical_team_name(m.get('team2')) if callable(getattr(bot, 'canonical_team_name', None)) else m.get('team2')
    t1 = t1 or m.get('team1') or ''
    t2 = t2 or m.get('team2') or ''
    if isinstance(obj, dict) and _numeric_score(obj):
        lines = [f"⚽ {t1} {obj.get('score1')} - {obj.get('score2')} {t2}", obj.get('status') or 'الحالة قيد التحديث', '']
        lines.extend(_goal_lines(obj))
        lines.append(f"المصدر: {obj.get('actual_source') or obj.get('source') or 'ESPN'}")
        lines.append(f"آخر تحديث: {now_makkah()}")
        return '\n'.join([x for x in lines if x is not None]).strip()
    return f"🏟️ {t1} × {t2}\n📌 الحالة: لم تبدأ / قيد التحديث\n⏰ الوقت: {m.get('time') or '-'}\nآخر تحديث: {now_makkah()}"


def _build_day_cache(d, force=False):
    d = _norm_date(d)
    debug = []
    events = _espn_events_for_day(d, force=force, debug=debug)
    fixtures = _fixtures_for_date(d)
    matches = []
    for m in fixtures:
        if callable(getattr(bot, '_has_unknown', None)):
            try:
                if bot._has_unknown(m):
                    continue
            except Exception:
                pass
        obj = None
        for ev in events:
            try:
                obj = _event_matches_fixture(ev, m)
                if obj:
                    break
            except Exception:
                continue
        mid = str(m.get('id') or f"{d}|{m.get('team1')}|{m.get('team2')}")
        matches.append({
            'id': mid,
            'date': d,
            'time': m.get('time'),
            'team1': bot.canonical_team_name(m.get('team1')) if callable(getattr(bot, 'canonical_team_name', None)) else m.get('team1'),
            'team2': bot.canonical_team_name(m.get('team2')) if callable(getattr(bot, 'canonical_team_name', None)) else m.get('team2'),
            'label': _label_for_match(m, obj),
            'detail': _detail_for_match(m, obj),
            'obj': obj or {},
        })
    return {'date': d, 'title': _fixture_title(d), 'events_count': len(events), 'matches': matches, 'debug': debug[-12:], 'updated_at': now_makkah()}


def _update_today_source():
    d = _active_live_date()
    # أول مرة اسمح بالـ force، بعدها اعتمد كاش ESPN لتبقى سريعة.
    first = not bool((load_cache().get('status') or {}).get('last_success'))
    payload = _build_day_cache(d, force=first)
    put_raw('live_matches', payload)
    put_raw('matches_today', payload)
    return payload


def _update_previous_results_source():
    # تحديث خفيف لنتائج الأيام السابقة: يستخدم كاش اليوم الموجود، ولا يجبر ESPN لكل الأيام.
    out = {}
    active = _active_live_date()
    for d in _fixture_dates_upto_active(limit=18):
        try:
            force = (_norm_date(d) == _norm_date(active))
            out[_norm_date(d)] = _build_day_cache(d, force=False if not force else False)
        except Exception as e:
            out[_norm_date(d)] = {'date': d, 'error': f'{type(e).__name__}: {str(e)[:120]}', 'matches': []}
    put_raw('results_by_date', out)
    return out


def _live_text_from_payload(payload=None):
    payload = payload or _v603_raw_value('live_matches', {})
    if not isinstance(payload, dict): payload = {}
    title = payload.get('title') or _fixture_title(payload.get('date') or _active_live_date())
    lines = [f'📺 مباشر الآن — {title}', f'آخر تحديث: {now_makkah()}', '', 'اختر مباراة من الأزرار لعرض التفاصيل.']
    matches = payload.get('matches') if isinstance(payload.get('matches'), list) else []
    if not matches:
        lines += ['', 'لا توجد مباريات في اليوم النشط حاليًا.']
    return '\n'.join(lines).strip()


def _v603_raw_value(key, default=None):
    try:
        raw = load_cache().get('raw') or {}
        item = raw.get(key) or {}
        if isinstance(item, dict) and 'value' in item:
            return item.get('value')
    except Exception:
        pass
    return default


def _groups_raw():
    groups = None
    for call in (
        lambda: bot.fetch_current_groups('latest'),
        lambda: bot.fetch_current_groups('official'),
        lambda: bot.fetch_standings_from_espn_live_results(),
    ):
        try:
            groups = call()
            if groups:
                break
        except Exception:
            continue
    if not groups:
        raise RuntimeError('لم يرجع ترتيب المجموعات من المصدر')
    return groups


def _groups_text(groups=None):
    groups = groups or _groups_raw()
    try:
        return bot.build_groups_text(groups, 'ESPN Live Results')
    except Exception:
        return '📊 ترتيب المجموعات\nتم تحديث بيانات الترتيب.\nآخر تحديث: ' + now_makkah()


def _board_text():
    for name in ('_V55_ORIG_BOARD_TEXT', '_v32_tournament_board_text', '_v50_board_text'):
        fn = getattr(bot, name, None)
        if callable(fn):
            try:
                txt = fn(True)
                if txt and 'جاري تجهيز' not in str(txt):
                    return str(txt)
            except Exception:
                continue
    return '🏆 لوحة البطولة\nلا توجد بيانات كافية حاليًا.'


def _top_scorers_text():
    items = []
    if callable(getattr(bot, 'fetch_espn_top_scorers', None)):
        try:
            # الدالة الأصلية في ملف ممتاز بدون باراميتر. لا نمرر force حتى لا يصير TypeError/قراءة غلط.
            items = bot.fetch_espn_top_scorers() or []
        except Exception:
            items = []
    if not items:
        return '🏆 هدافين البطولة\nلا توجد بيانات هدافين جاهزة حاليًا.'
    normalized = []
    for it in items[:50]:
        if isinstance(it, dict):
            name = it.get('name') or it.get('player') or it.get('athlete') or '-'
            team = it.get('team') or it.get('country') or ''
            goals = _safe_int(it.get('goals') or it.get('goal') or it.get('totalGoals') or it.get('value'))
        else:
            try:
                name, goals, team = it[0], it[1], (it[2] if len(it) > 2 else '')
                goals = _safe_int(goals)
            except Exception:
                continue
        normalized.append({'name': str(name), 'team': str(team or ''), 'goals': goals})
    normalized.sort(key=lambda x: (-x['goals'], x['name']))
    lines = ['🏆 هدافين البطولة', f'آخر تحديث: {now_makkah()}', '']
    for i, it in enumerate(normalized[:20], start=1):
        team = f" — {it['team']}" if it.get('team') else ''
        goals = it.get('goals') or 0
        word = 'هدف' if goals == 1 else 'أهداف'
        lines.append(f"{i}. {it['name']}{team} — {goals} {word}")
    return '\n'.join(lines).strip()


def _call_text(name, fallback):
    fn = getattr(bot, name, None)
    if not callable(fn):
        return fallback
    # بعض دوال ملف ممتاز تستقبل force وبعضها لا. نحافظ على الاثنين بدون كسر.
    try:
        return fn(True)
    except TypeError:
        return fn()


async def refresh_stage(key, title, pct, fn, timeout, block=True, raw=False):
    set_status(state='running', progress=pct, stage=title, next_expected=next_time())
    try:
        result = await run_sync_with_timeout(fn, timeout=timeout)
        if raw:
            put_raw(key, result)
            text = f'{title}\nتم تحديث البيانات الخام.\nآخر تحديث: {now_makkah()}'
        else:
            text = str(result or '').strip()
        if block and text:
            put_block(key, text, {'stage': title, 'ok': True})
        return True
    except Exception as e:
        old = existing_text(key)
        if block and not old:
            put_block(key, f'{title}\n⚠️ لم يجهز هذا القسم بعد.\nسيحاول التحديث التلقائي مرة أخرى.\nالسبب: {type(e).__name__}: {str(e)[:180]}', {'stage': title, 'ok': False, 'error': str(e)[:250]})
        set_status(last_error=f'{title}: {type(e).__name__}: {str(e)[:180]}')
        return False


async def refresh_all_once(first_run=False):
    timeout = FIRST_STAGE_TIMEOUT if first_run else STAGE_TIMEOUT
    set_status(state='running', progress=1, stage='بدء التحديث', started_at=now_makkah(), next_expected=next_time(), last_error='')
    groups_holder = {'groups': None}
    today_holder = {'payload': None}

    def source_today():
        today_holder['payload'] = _update_today_source()
        return today_holder['payload']
    def source_previous():
        return _update_previous_results_source()
    def live_block():
        return _live_text_from_payload(today_holder.get('payload') or _v603_raw_value('live_matches', {}))
    def raw_groups():
        groups_holder['groups'] = _groups_raw()
        return groups_holder['groups']
    def groups_block():
        return _groups_text(groups_holder.get('groups'))

    # ترتيب مهم: نفس طريقة ملف ممتاز: بيانات ESPN الخام أولًا، ثم الكاشات، ثم الحسابات.
    # لا توجد مهلة خارجية توقف المرحلة؛ الدوال الأصلية هي التي تتحكم بمهل طلبات ESPN.
    stages = [
        ('matches_today', '📡 سحب مباريات اليوم من ESPN', 8, source_today, False, True),
        ('results_by_date', '📡 سحب نتائج مباريات البطولة', 16, source_previous, False, True),
        ('raw_groups', '📡 سحب ترتيب المجموعات من ESPN', 24, raw_groups, False, True),
        ('live', '📺 مباشر الآن', 32, live_block, True, False),
        ('groups', '📊 ترتيب المجموعات', 40, groups_block, True, False),
        ('scorers', '🏆 هدافين البطولة', 50, _top_scorers_text, True, False),
        ('board', '🏆 لوحة البطولة', 60, _board_text, True, False),
        ('thirds', '🥉 أفضل الثوالث الآن', 70, lambda: _call_text('_v32_best_thirds_text', '🥉 أفضل الثوالث الآن\nلا توجد بيانات كافية.'), True, False),
        ('qualified', '✅ المتأهلون رسميًا', 78, lambda: _call_text('_v56_qualified_text', '✅ المتأهلون رسميًا\nلا توجد بيانات كافية.'), True, False),
        ('eliminated', '❌ المغادرون رسميًا', 84, lambda: bot._v57_eliminated_text() if hasattr(bot, '_v57_eliminated_text') else '❌ المغادرون رسميًا\nلا توجد بيانات كافية.', True, False),
        ('r32', '🏟️ مباريات دور الـ32', 92, lambda: _call_text('_v56_round32_text', '🏟️ مباريات دور الـ32\nلم تتحدد بعد.'), True, False),
        ('watch', '👀 وش أتابع الجولة الأخيرة؟', 97, lambda: _call_text('_v56_watch_last_round_text', '👀 وش أتابع الجولة الأخيرة؟\nلا توجد بيانات كافية.'), True, False),
    ]
    ok_count = 0
    for key, title, pct, fn, block, raw in stages:
        if await refresh_stage(key, title, pct, fn, timeout=timeout, block=block, raw=raw):
            ok_count += 1
    set_status(state='ready', progress=100, stage='اكتمل التحديث', last_success=now_makkah(), next_expected=next_time(), ok_sections=ok_count, total_sections=len(stages))


async def main_loop():
    print('Masif old updater V60.4 started. interval=', INTERVAL, 'stage_timeout=', STAGE_TIMEOUT, 'first_timeout=', FIRST_STAGE_TIMEOUT, flush=True)
    first_run = not bool((load_cache().get('status') or {}).get('last_success'))
    while True:
        try:
            if os.path.exists(LOCK_FILE):
                age = time.time() - os.path.getmtime(LOCK_FILE)
                if age < max(INTERVAL * 2, FIRST_STAGE_TIMEOUT + 60):
                    data = load_cache(); st = data.setdefault('status', {})
                    st['skipped_count'] = int(st.get('skipped_count') or 0) + 1
                    st['next_expected'] = next_time()
                    save_cache(data)
                    await asyncio.sleep(INTERVAL)
                    continue
            with open(LOCK_FILE, 'w') as f:
                f.write(str(time.time()))
            await refresh_all_once(first_run=first_run)
            first_run = False
        except Exception:
            set_status(state='error', progress=0, stage='خطأ في خدمة التحديث', last_error=traceback.format_exc()[-900:], next_expected=next_time())
        finally:
            try:
                if os.path.exists(LOCK_FILE): os.remove(LOCK_FILE)
            except Exception:
                pass
        await asyncio.sleep(INTERVAL)


if __name__ == '__main__':
    asyncio.run(main_loop())
