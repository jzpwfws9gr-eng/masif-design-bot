#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""V60.8 lightweight updater.
No ESPN pull, no cache build. Bot pulls live/results/scorers directly مثل ملف ممتاز.
"""
from datetime import datetime, timedelta
import json, time, os
STATUS_FILE = os.environ.get("MASIF_UPDATE_STATUS_FILE", "update_status.json")
while True:
    now = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
    data = {
        "state": "ready",
        "percent": 100,
        "stage": "السحب المباشر يعمل بطريقة ملف ممتاز عند الضغط",
        "heartbeat": now,
        "note": "لا يوجد كاش ولا تحديث ESPN بالخلفية في V60.8"
    }
    try:
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    time.sleep(60)
