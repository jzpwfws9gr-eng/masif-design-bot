import os, json, time, asyncio, traceback
from datetime import datetime, timedelta

# خدمة التحديث المنفصلة: السحب الحقيقي يتم هنا، وبوت المستخدمين يقرأ الكاش فقط.
os.environ.setdefault('MASIF_USER_BOT_CACHE_ONLY', '0')
import bot

CACHE_FILE = os.environ.get('MASIF_SHARED_CACHE_FILE', 'masif_shared_cache.json')
INTERVAL = int(os.environ.get('MASIF_UPDATE_INTERVAL_SECONDS', '120'))
LOCK_FILE = os.environ.get('MASIF_UPDATE_LOCK_FILE', 'masif_update.lock')
FIRST_STAGE_TIMEOUT = int(os.environ.get('MASIF_FIRST_STAGE_TIMEOUT_SECONDS', '120'))
STAGE_TIMEOUT = int(os.environ.get('MASIF_STAGE_TIMEOUT_SECONDS', '120'))


def now_makkah():
    return (datetime.utcnow() + timedelta(hours=3)).strftime('%Y-%m-%d %H:%M:%S')


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


async def run_sync_with_timeout(func, timeout):
    return await asyncio.wait_for(asyncio.to_thread(func), timeout=int(timeout))


def _safe_int(x, default=0):
    try: return int(float(str(x).strip()))
    except Exception: return default


def _active_live_date():
    for name in ('_v41_active_live_date', '_v40_active_live_date', '_v33_active_live_date', '_v29_active_live_date'):
        fn = getattr(bot, name, None)
        if callable(fn):
            try:
                v = fn()
                if v: return v
            except Exception:
                pass
    try:
        return (datetime.utcnow() + timedelta(hours=3)).strftime('%d/%m/%Y')
    except Exception:
        return ''


def _fixture_title(d):
    try:
        return bot._v26_fixture_title(d)
    except Exception:
        return str(d or '')


def _update_live_raw():
    """يسحب مباشر/نتائج اليوم النشط من ESPN بنفس مسار النسخة القديمة."""
    d = _active_live_date()
    updated, debug = 0, []
    if d and callable(getattr(bot, '_v41_update_day_results', None)):
        try:
            updated, debug = bot._v41_update_day_results(d, True)
        except Exception as e:
            debug = [f'{type(e).__name__}: {str(e)[:180]}']
            raise
    return {'date': d, 'updated': updated, 'debug': debug[-10:] if isinstance(debug, list) else []}


def _live_text_from_cache():
    d = _active_live_date()
    title = _fixture_title(d)
    rows = []
    try:
        matches = bot._v40_live_rows(d) if callable(getattr(bot, '_v40_live_rows', None)) else []
    except Exception:
        matches = []
    for m in matches or []:
        try:
            label = bot._v40_live_button_label(m) if callable(getattr(bot, '_v40_live_button_label', None)) else ''
        except Exception:
            label = ''
        if not label:
            t1 = getattr(bot, 'canonical_team_name', lambda x: x)(m.get('team1')) or m.get('team1') or ''
            t2 = getattr(bot, 'canonical_team_name', lambda x: x)(m.get('team2')) or m.get('team2') or ''
            label = f'{t1} × {t2}'
        rows.append('• ' + str(label))
    if not rows:
        rows.append('لا توجد مباريات في مباشر الآن حسب جدول البطولة.')
    return '\n'.join([f'📺 مباشر الآن — {title}', f'آخر تحديث: {now_makkah()}', '', *rows]).strip()


def _groups_raw():
    """الترتيب الحقيقي من المصدر المعتمد في النسخة القديمة، غالبًا ESPN Live Results."""
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
        lines = ['📊 ترتيب المجموعات', f'آخر تحديث: {now_makkah()}', '']
        if isinstance(groups, dict):
            iterator = groups.items()
        else:
            iterator = groups or []
        for item in iterator:
            try:
                code, rows = item if isinstance(item, tuple) else (item.get('group'), item.get('rows'))
                lines.append(f'المجموعة {code}')
                for r in rows or []:
                    team = r.get('team') or r.get('name') or '-'
                    pts = r.get('Pts') or r.get('pts') or r.get('points') or 0
                    gd = r.get('GD') or r.get('gd') or 0
                    lines.append(f'- {team}: {pts} نقطة | فارق {gd}')
                lines.append('')
            except Exception:
                pass
        return '\n'.join(lines).strip()


def _board_text():
    # يقرأ من نفس دوال اللوحة الحقيقية، مع force حتى يسحب من البيانات الجديدة.
    for name in ('_V55_ORIG_BOARD_TEXT', '_v32_tournament_board_text', '_v50_board_text'):
        fn = getattr(bot, name, None)
        if callable(fn):
            txt = fn(True)
            if txt and 'جاري تجهيز' not in str(txt):
                return txt
    return '🏆 لوحة البطولة\nلا توجد بيانات كافية حاليًا.'


def _top_scorers_text():
    items = []
    if callable(getattr(bot, 'fetch_espn_top_scorers', None)):
        items = bot.fetch_espn_top_scorers(True) or []
    if not items:
        return '🏆 هدافين البطولة\nلا توجد بيانات هدافين جاهزة حاليًا.'
    normalized = []
    for it in items[:30]:
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
    return fn(True)


async def refresh_stage(key, title, pct, fn, timeout):
    set_status(state='running', progress=pct, stage=title, next_expected=next_time())
    try:
        result = await run_sync_with_timeout(fn, timeout=timeout)
        if isinstance(result, (dict, list)):
            put_raw(key, result)
            text = f'{title}\nتم تحديث البيانات الخام.\nآخر تحديث: {now_makkah()}'
        else:
            text = str(result or '').strip()
        if text:
            put_block(key, text, {'stage': title, 'ok': True})
        return True
    except Exception as e:
        old = existing_text(key)
        if not old:
            put_block(key, f'{title}\n⚠️ لم يجهز هذا القسم بعد.\nسيحاول التحديث التلقائي مرة أخرى.\nالسبب: {type(e).__name__}: {str(e)[:180]}', {'stage': title, 'ok': False, 'error': str(e)[:250]})
        set_status(last_error=f'{title}: {type(e).__name__}: {str(e)[:180]}')
        return False


async def refresh_all_once(first_run=False):
    timeout = FIRST_STAGE_TIMEOUT if first_run else STAGE_TIMEOUT
    set_status(state='running', progress=1, stage='بدء التحديث', started_at=now_makkah(), next_expected=next_time(), last_error='')

    # الترتيب مهم: الخام أولًا، ثم الأقسام التي تعتمد عليه.
    groups_holder = {'groups': None}

    def raw_live():
        return _update_live_raw()

    def raw_groups():
        groups_holder['groups'] = _groups_raw()
        return groups_holder['groups']

    def groups_block():
        return _groups_text(groups_holder.get('groups'))

    stages = [
        ('raw_live', '📡 سحب مباشر الآن من ESPN', 8, raw_live),
        ('raw_groups', '📡 سحب ترتيب المجموعات من ESPN', 18, raw_groups),
        ('groups', '📊 ترتيب المجموعات', 25, groups_block),
        ('live', '📺 مباشر الآن', 32, _live_text_from_cache),
        ('scorers', '🏆 هدافين البطولة', 40, _top_scorers_text),
        ('board', '🏆 لوحة البطولة', 50, _board_text),
        ('thirds', '🥉 أفضل الثوالث الآن', 62, lambda: _call_text('_v32_best_thirds_text', '🥉 أفضل الثوالث الآن\nلا توجد بيانات كافية.')),
        ('qualified', '✅ المتأهلون رسميًا', 72, lambda: _call_text('_v56_qualified_text', '✅ المتأهلون رسميًا\nلا توجد بيانات كافية.')),
        ('eliminated', '❌ المغادرون رسميًا', 80, lambda: bot._v57_eliminated_text() if hasattr(bot, '_v57_eliminated_text') else '❌ المغادرون رسميًا\nلا توجد بيانات كافية.'),
        ('r32', '🏟️ مباريات دور الـ32', 90, lambda: _call_text('_v56_round32_text', '🏟️ مباريات دور الـ32\nلم تتحدد بعد.')),
        ('watch', '👀 وش أتابع الجولة الأخيرة؟', 96, lambda: _call_text('_v56_watch_last_round_text', '👀 وش أتابع الجولة الأخيرة؟\nلا توجد بيانات كافية.')),
    ]
    ok_count = 0
    for key, title, pct, fn in stages:
        if await refresh_stage(key, title, pct, fn, timeout=timeout):
            ok_count += 1
    set_status(state='ready', progress=100, stage='اكتمل التحديث', last_success=now_makkah(), next_expected=next_time(), ok_sections=ok_count, total_sections=len(stages))


async def main_loop():
    print('Masif old updater V60.2 started. interval=', INTERVAL, 'stage_timeout=', STAGE_TIMEOUT, 'first_timeout=', FIRST_STAGE_TIMEOUT)
    first_run = not bool((load_cache().get('status') or {}).get('last_success'))
    while True:
        try:
            if os.path.exists(LOCK_FILE):
                age = time.time() - os.path.getmtime(LOCK_FILE)
                # لا نسمح بتداخل تحديثين. لو القفل قديم جدًا نكسره.
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
