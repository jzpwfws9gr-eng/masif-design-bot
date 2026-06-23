import os, json, time, asyncio, traceback
from datetime import datetime, timedelta

os.environ.setdefault('MASIF_USER_BOT_CACHE_ONLY', '0')
import bot

CACHE_FILE = os.environ.get('MASIF_SHARED_CACHE_FILE', 'masif_shared_cache.json')
INTERVAL = int(os.environ.get('MASIF_UPDATE_INTERVAL_SECONDS', '120'))
LOCK_FILE = os.environ.get('MASIF_UPDATE_LOCK_FILE', 'masif_update.lock')


def now_makkah():
    return (datetime.utcnow() + timedelta(hours=3)).strftime('%Y-%m-%d %H:%M:%S')


def next_time():
    return (datetime.utcnow() + timedelta(hours=3, seconds=INTERVAL)).strftime('%H:%M:%S')


def load_cache():
    try:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
    except Exception:
        pass
    return {'blocks': {}, 'status': {}}


def save_cache(data):
    tmp = CACHE_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data or {}, f, ensure_ascii=False, indent=2)
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


async def run_sync_with_timeout(func, timeout=75):
    return await asyncio.wait_for(asyncio.to_thread(func), timeout=timeout)


async def refresh_all_once():
    set_status(state='running', progress=1, stage='بدء التحديث', started_at=now_makkah(), next_expected=next_time(), last_error='')
    stages = [
        ('live', '📺 مباشر الآن', 10, lambda: getattr(bot, '_v49_live_now_text', lambda *a, **k: bot._v60_cache_text('live','📺 مباشر الآن'))()),
        ('board', '🏆 لوحة البطولة', 25, lambda: bot._v562_board_text_for_user() if hasattr(bot, '_v562_board_text_for_user') else bot._v50_board_text(True)),
        ('thirds', '🥉 أفضل الثوالث الآن', 40, lambda: bot._v32_best_thirds_text(True)),
        ('qualified', '✅ المتأهلون رسميًا', 55, lambda: bot._v56_qualified_text(True)),
        ('eliminated', '❌ المغادرون رسميًا', 65, lambda: bot._v57_eliminated_text() if hasattr(bot, '_v57_eliminated_text') else '❌ المغادرون رسميًا\nلا يوجد كاش.'),
        ('r32', '🏟️ مباريات دور الـ32', 75, lambda: bot._v56_round32_text(True)),
        ('watch', '👀 وش أتابع الجولة الأخيرة؟', 85, lambda: bot._v56_watch_last_round_text(True)),
        ('scorers', '🏆 هدافين البطولة', 95, lambda: bot.build_top_scorers_caption(bot.fetch_espn_top_scorers(True)) if hasattr(bot, 'fetch_espn_top_scorers') and hasattr(bot, 'build_top_scorers_caption') else '🏆 هدافين البطولة\nغير متوفر حاليًا.'),
    ]
    for key, title, pct, fn in stages:
        set_status(state='running', progress=pct, stage=title, next_expected=next_time())
        try:
            text = await run_sync_with_timeout(fn, timeout=int(os.environ.get('MASIF_STAGE_TIMEOUT_SECONDS', '75')))
            put_block(key, text or f'{title}\nلا توجد بيانات حالية.', {'stage': title})
        except Exception as e:
            data = load_cache()
            old = ((data.get('blocks') or {}).get(key) or {}).get('text')
            if not old:
                put_block(key, f'{title}\n⚠️ تعذر تجهيز هذا القسم حاليًا.\nالسبب: {type(e).__name__}: {str(e)[:180]}', {'error': str(e)[:250]})
            set_status(last_error=f'{title}: {type(e).__name__}: {str(e)[:180]}')
    set_status(state='ready', progress=100, stage='اكتمل التحديث', last_success=now_makkah(), next_expected=next_time())


async def main_loop():
    print('Masif old updater started. interval=', INTERVAL)
    while True:
        try:
            if os.path.exists(LOCK_FILE):
                age = time.time() - os.path.getmtime(LOCK_FILE)
                if age < max(INTERVAL, 180):
                    data = load_cache(); st = data.setdefault('status', {})
                    st['skipped_count'] = int(st.get('skipped_count') or 0) + 1
                    st['next_expected'] = next_time()
                    save_cache(data)
                    await asyncio.sleep(INTERVAL)
                    continue
            with open(LOCK_FILE, 'w') as f: f.write(str(time.time()))
            await refresh_all_once()
        except Exception as e:
            set_status(state='error', progress=0, stage='خطأ في خدمة التحديث', last_error=traceback.format_exc()[-900:], next_expected=next_time())
        finally:
            try:
                if os.path.exists(LOCK_FILE): os.remove(LOCK_FILE)
            except Exception:
                pass
        await asyncio.sleep(INTERVAL)


if __name__ == '__main__':
    asyncio.run(main_loop())
