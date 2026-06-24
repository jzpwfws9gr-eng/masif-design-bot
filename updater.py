import json, time, os
from datetime import datetime, timedelta

STATUS_FILE = os.getenv("MASIF_SHARED_CACHE_FILE", "masif_shared_cache.json")

def write_status():
    try:
        now = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
        data = {
            "update_status": {
                "state": "ready",
                "progress": 100,
                "stage": "السحب المباشر القديم مفعل داخل bot.py عند الضغط",
                "heartbeat": now,
                "note": "لا كاش جديد ولا تحديث ثقيل؛ bot.py يعمل بطريقة ملف ممتاز."
            }
        }
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

if __name__ == "__main__":
    while True:
        write_status()
        time.sleep(60)
