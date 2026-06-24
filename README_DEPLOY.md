# Masif Old Improved V60.4 — Restore Excellent Pull Routes

## التشغيل في Railway
Start Command:

```bash
sh -c "python updater.py & python bot.py"
```

## المطلوب رفعه إلى GitHub
- bot.py
- updater.py
- README_DEPLOY.md

## المتغيرات
- BOT_TOKEN
- ADMIN_IDS
- MASIF_SHARED_CACHE_FILE=masif_shared_cache.json
- MASIF_UPDATE_INTERVAL_SECONDS=120

## أهم تغييرات V60.4
- رجوع طريقة السحب الأولى من ملف "ممتتتاز .py" عبر دوال bot.py الأصلية.
- إلغاء الـ Timeout الخارجي الذي كان يوقف كل مرحلة بعد دقيقتين.
- تبقى مهلة طلبات ESPN الداخلية كما في الكود القديم، حتى لا يعلق الطلب الواحد للأبد.
- الهدافين يتم سحبهم بالطريقة القديمة بدون تمرير force خاطئ.
- مباشر الآن يرجع أزرار + تفاصيل من الكاش.
- بوت المستخدمين يقرأ الكاش فقط ولا يسحب من ESPN وقت الضغط.
- التحديث التلقائي لم يُلغَ.
- احتمالات المغادرة وسيناريوهات الخروج محذوفة.
