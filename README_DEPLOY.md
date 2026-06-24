# Masif Old Improved V60.3

تشغيل Railway:

```bash
sh -c "python updater.py & python bot.py"
```

المتغيرات المطلوبة:
- BOT_TOKEN
- ADMIN_IDS أو ADMIN_ID
- MASIF_SHARED_CACHE_FILE=masif_shared_cache.json
- MASIF_UPDATE_INTERVAL_SECONDS=120
- MASIF_FIRST_STAGE_TIMEOUT_SECONDS=120
- MASIF_STAGE_TIMEOUT_SECONDS=120

V60.3:
- استرجاع طريقة القديم في مباشر الآن: أزرار مباريات ثم تفاصيل.
- السحب الحقيقي من ESPN في updater.py.
- بوت المستخدمين يقرأ من الكاش فقط.
- حذف احتمالات/سيناريوهات الخروج.
