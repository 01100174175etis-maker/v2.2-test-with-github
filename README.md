# النظام المحاسبي (v2.2) - نسخة ويب (Streamlit)

تطبيق ويب بسيط لتسجيل وعرض أصناف العملاء وتجميع ملف Excel رئيسي.

تشغيل محلي

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

3. شغّل الواجهة:

```bash
streamlit run streamlit_app.py
```

ملاحظات
- التطبيق يقرأ ويكتب ملفات Excel لكل عميل في مجلد المشروع نفسه.
- تأكد من إغلاق أي ملفات Excel مفتوحة قبل الحفظ.
- يمكنك تنزيل الملف الرئيسي (Master Excel) مباشرة من الواجهة.
