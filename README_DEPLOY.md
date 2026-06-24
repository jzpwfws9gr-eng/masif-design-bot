# Masif Old Improved V60.2

## تشغيل Railway
Start Command:

```bash
sh -c "python updater.py & python bot.py"
```

## Variables
- BOT_TOKEN
- ADMIN_IDS=678392605
- MASIF_SHARED_CACHE_FILE=masif_shared_cache.json
- MASIF_UPDATE_INTERVAL_SECONDS=120
- MASIF_FIRST_STAGE_TIMEOUT_SECONDS=120
- MASIF_STAGE_TIMEOUT_SECONDS=120

## V60.2
- خدمة التحديث تسحب البيانات الحقيقية من ESPN/المصادر الأصلية وتحفظ الكاش.
- بوت المستخدمين يقرأ من الكاش فقط ولا يسحب عند الضغط.
- أول تحديث يعطي كل قسم مهلة دقيقتين.
- التحديث يمشي بترتيب ثابت: خام ESPN ثم الأقسام الخفيفة ثم الحسابات.
- التحديث التلقائي مستمر ولم يتم إلغاؤه.
- تم حذف احتمالات المغادرة/سيناريوهات الخروج بالكامل من هذه النسخة.
