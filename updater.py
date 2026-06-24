import os, json, time, asyncio
from datetime import datetime, timedelta

# خدمة التحديث V60.5
# رجعنا السحب لطريقة ملف ممتاز داخل bot.py عند الضغط.
# updater هنا نبض حالة فقط حتى لا يعلق على ESPN ولا يمنع البوت.

CACHE_FILE = os.environ.get('MASIF_SHARED_CACHE_FILE', 'masif_shared_cache.json')
INTERVAL = int(os.environ.get('MASIF_UPDATE_INTERVAL_SECONDS', '120'))


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
    return {}


def save_cache(data):
    data = data or {}
    data.setdefault('blocks', {})
    data.setdefault('raw', {})
    data.setdefault('status', {})
    tmp = CACHE_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CACHE_FILE)


def set_ready_status():
    data = load_cache()
    st = data.setdefault('status', {})
    st.update({
        'state': 'ready',
        'progress': 100,
        'stage': 'السحب المباشر القديم مفعل عند الضغط',
        'started_at': st.get('started_at') or now_makkah(),
        'last_success': now_makkah(),
        'next_expected': next_time(),
        'last_error': '',
        'mode': 'direct_old_pull_v605',
    })
    save_cache(data)


async def main_loop():
    print('Masif updater V60.5 heartbeat started — direct old pull is handled by bot.py', flush=True)
    while True:
        try:
            set_ready_status()
        except Exception as e:
            print('heartbeat error:', repr(e), flush=True)
        await asyncio.sleep(INTERVAL)


if __name__ == '__main__':
    asyncio.run(main_loop())
