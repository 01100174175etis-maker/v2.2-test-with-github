# النظام المحاسبي (v2.2) - نسخة ويب (Streamlit)

تغييرات مهمة في هذه النسخة:
- الانتقال لاستخدام قاعدة بيانات محلية SQLite (data.db) لتخزين السجلات بدلاً من الكتابة المباشرة إلى ملفات Excel.
  - عند التشغيل لأول مرة، إن وُجدت ملفات Excel للعميل فسيتم استيرادها تلقائيًا إلى قاعدة البيانات.
- إضافة مصادقة بسيطة اعتمادًا على متغير البيئة APP_PASSWORD.
- إمكانية إنشاء ملف Excel رئيسي وتجميع الأوراق من قاعدة البيانات (مع دعم ملفات Excel القديمة).
- إضافة Dockerfile و docker-compose لتشغيل التطبيق بسهولة دون إعداد بيئة محلية مفصّلة.

متطلبات

- Python 3.8+

التثبيت والتشغيل محليًا

1. أنشئ وفعّل بيئة افتراضية (اختياري):

```bash
python -m venv venv
# Linux/macOS
source venv/bin/activate
# Windows
venv\Scripts\activate
```

2. ثبّت المتطلبات:

```bash
pip install -r requirements.txt
```

3. تشغيل التطبيق:

```bash
# تشغيل في وضع headless (مناسب للخوادم):
streamlit run streamlit_app.py --server.headless=true

# أو بشكل عادي:
streamlit run streamlit_app.py
```

تعيين كلمة المرور (اختياري للحماية)

- لتفعيل حماية بسيطة واجهة الويب اضبط متغير البيئة APP_PASSWORD قبل تشغيل التطبيق. مثال (Linux/macOS):

```bash
export APP_PASSWORD="yourpassword"
streamlit run streamlit_app.py --server.headless=true
```

Windows PowerShell:

```powershell
$env:APP_PASSWORD = "yourpassword"
streamlit run streamlit_app.py --server.headless=true
```

Docker

1. اخلق ملف `.env` بناءً على `.env.example` وضع APP_PASSWORD إن رغبت بالحماية.
2. شغّل باستخدام docker-compose:

```bash
docker-compose up --build
```

الواجهه ستكون متاحة عادة على http://localhost:8501

ملاحظات

- قاعدة البيانات: data.db (مُدرَجة في .gitignore تلقائيًا)
- تأكد من إغلاق أي ملفات Excel قبل تشغيل التطبيق عند وجودها في المجلد.
