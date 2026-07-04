#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""لا يوجد تحديث خارجي في نسخة المسابقة اليدوية."""
import time, json, os
from datetime import datetime, timedelta
STATUS_FILE = os.environ.get("MASIF_UPDATE_STATUS_FILE", "update_status.json")
while True:
    now = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump({"state":"ready","percent":100,"stage":"مسابقة يدوية فقط","heartbeat":now,"note":"لا يوجد سحب مواقع"}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    time.sleep(60)
