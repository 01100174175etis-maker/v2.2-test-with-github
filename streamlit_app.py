import os
import glob
import re
from io import BytesIO
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np
from sqlalchemy import create_engine, Column, Integer, String, Float
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from dotenv import load_dotenv

# Load .env if present (development convenience)
load_dotenv()

# Configuration
ITEMS_FILE = "items_list.txt"
MASTER_FILE = "الملف_الرئيسي_المحاسبي.xlsx"
DB_FILE = "data.db"
DB_URL = f"sqlite:///{DB_FILE}"
APP_PASSWORD = os.environ.get("APP_PASSWORD")  # إذا وُضع، ستُطلب كلمة المرور للوصول

st.set_page_config(page_title="النظام المحاسبي - ويب", layout="centered")
st.title("النظام المحاسبي (نسخة ويب) — v2.2")

# -------------------- قاعدة البيانات (SQLite عبر SQLAlchemy) --------------------
Base = declarative_base()

class Entry(Base):
    __tablename__ = "entries"
    id = Column(Integer, primary_key=True, index=True)
    customer = Column(String, index=True)
    item_type = Column(String)
    qty = Column(Integer)
    price = Column(Float)
    total = Column(Float)
    date = Column(String)  # محفوظ كنص بصيغة YYYY-MM-DD


engine = create_engine(DB_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


def init_db():
    Base.metadata.create_all(bind=engine)

    # إذا كانت قاعدة البيانات فارغة ولديك ملفات Excel قديمة، نحاول استيرادها تلقائياً
    session = SessionLocal()
    try:
        count = session.query(Entry).count()
        if count == 0:
            # استعِد ملفات Excel الموجودة واستوردها
            files = glob.glob("*.xlsx")
            valid = [f for f in files if not f.startswith("~$") and os.path.basename(f) != MASTER_FILE]
            for f in valid:
                name = os.path.splitext(os.path.basename(f))[0]
                try:
                    df = pd.read_excel(f, engine='openpyxl')
                except Exception:
                    df = pd.DataFrame()
                if not df.empty:
                    # نتوقع أعمدة: التاريخ, الإجمالي, السعر, الكمية, نوع الصنف
                    for _, row in df.iterrows():
                        try:
                            date_str = pd.to_datetime(row.get("التاريخ", ""), errors='coerce')
                            if pd.isna(date_str):
                                date_txt = str(row.get("التاريخ", ""))
                            else:
                                date_txt = date_str.strftime('%Y-%m-%d')
                        except Exception:
                            date_txt = str(row.get("التاريخ", ""))
                        try:
                            qty = int(row.get("الكمية", 0))
                        except Exception:
                            qty = 0
                        try:
                            price = float(row.get("السعر", 0))
                        except Exception:
                            price = 0.0
                        total = qty * price
                        entry = Entry(customer=name, item_type=str(row.get("نوع الصنف", "")), qty=qty, price=price, total=total, date=date_txt)
                        session.add(entry)
            session.commit()
    finally:
        session.close()


# -------------------- مساعدة الملفات القديمة --------------------

def sanitize_sheet_name(name: str) -> str:
    if not isinstance(name, str):
        name = str(name)
    cleaned = re.sub(r'[:\\/*?\[\]]', '_', name)
    cleaned = cleaned.strip()
    if len(cleaned) > 31:
        cleaned = cleaned[:31]
    if not cleaned:
        cleaned = "sheet"
    return cleaned


# -------------------- دوال تطبيقية (DB) --------------------

def add_entry_db(customer: str, item_type: str, qty: int, price: float, date_str: str):
    session = SessionLocal()
    try:
        total = qty * price
        entry = Entry(customer=customer, item_type=item_type, qty=qty, price=price, total=total, date=date_str)
        session.add(entry)
        session.commit()
        return True
    except Exception as e:
        st.error(f"خطأ عند إضافة السجل: {e}")
        session.rollback()
        return False
    finally:
        session.close()


def get_customers_db():
    session = SessionLocal()
    try:
        rows = session.query(Entry.customer).distinct().all()
        customers = [r[0] for r in rows if r[0]]
    finally:
        session.close()
    # اجمع بين العملاء الموجودين في DB وملفات Excel القديمة
    files = glob.glob("*.xlsx")
    files_customers = [os.path.splitext(os.path.basename(f))[0] for f in files if not f.startswith("~$") and os.path.basename(f) != MASTER_FILE]
    return sorted(set(customers + files_customers))


def get_entries_by_customer(customer: str):
    session = SessionLocal()
    try:
        rows = session.query(Entry).filter(Entry.customer == customer).order_by(Entry.id).all()
        df = pd.DataFrame([{
            "id": r.id,
            "التاريخ": r.date,
            "نوع الصنف": r.item_type,
            "الكمية": r.qty,
            "السعر": r.price,
            "الإجمالي": r.total
        } for r in rows])
    finally:
        session.close()

    # بالإضافة: إن كان هناك ملف Excel قديم ولم تُستورد بعض السجلات (نادر بعد الاستيراد) نقرأه كنسخة احتياط
    file_name = f"{customer}.xlsx"
    if os.path.exists(file_name):
        try:
            df_file = pd.read_excel(file_name, engine='openpyxl')
            # إذا كان df فارغ أو لم يُستورد، فلن يكون هناك تكرار (نحن قمنا باستيراد كل شيء في init_db إذا DB كانت فارغة سابقاً)
            # لكن لسلامة الأمور، إذا كانت قاعدة البيانات لا تحتوي أي صفوف نُظهر محتوى الملف
            session = SessionLocal()
            try:
                cnt = session.query(Entry).filter(Entry.customer == customer).count()
            finally:
                session.close()
            if cnt == 0 and not df_file.empty:
                # تحويل أعمدة لعرض موحد
                df_file = df_file.fillna("")
                display_df = pd.DataFrame()
                display_df["id"] = range(1, len(df_file) + 1)
                display_df["التاريخ"] = df_file.get("التاريخ", "")
                display_df["نوع الصنف"] = df_file.get("نوع الصنف", "")
                display_df["الكمية"] = df_file.get("الكمية", "")
                display_df["السعر"] = df_file.get("السعر", "")
                display_df["الإجمالي"] = df_file.get("الإجمالي", "")
                return display_df
        except Exception:
            pass

    if df.empty:
        return pd.DataFrame(columns=["id", "التاريخ", "نوع الصنف", "الكمية", "السعر", "الإجمالي"]) if 'df' in locals() else pd.DataFrame(columns=["id", "التاريخ", "نوع الصنف", "الكمية", "السعر", "الإجمالي"]) 
    return df


def delete_entry_db(entry_id: int):
    session = SessionLocal()
    try:
        row = session.query(Entry).filter(Entry.id == entry_id).first()
        if not row:
            return False
        session.delete(row)
        session.commit()
        return True
    except Exception as e:
        st.error(f"خطأ عند الحذف من DB: {e}")
        session.rollback()
        return False
    finally:
        session.close()


def make_master_from_db_in_memory():
    session = SessionLocal()
    try:
        customers = [r[0] for r in session.query(Entry.customer).distinct().all()]
        # أيضًا نأخذ العملاء من ملفات excel المُتبقية
        files = glob.glob("*.xlsx")
        files_customers = [os.path.splitext(os.path.basename(f))[0] for f in files if not f.startswith("~$") and os.path.basename(f) != MASTER_FILE]
        all_customers = sorted(set(customers + files_customers))

        buffer = BytesIO()
        summary = []
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            for cust in all_customers:
                # جلب من DB
                rows = session.query(Entry).filter(Entry.customer == cust).order_by(Entry.id).all()
                if rows:
                    df = pd.DataFrame([{
                        "التاريخ": r.date,
                        "نوع الصنف": r.item_type,
                        "الكمية": r.qty,
                        "السعر": r.price,
                        "الإجمالي": r.total
                    } for r in rows])
                else:
                    # جرب قراءة ملف العميل إن وُجد
                    file_name = f"{cust}.xlsx"
                    try:
                        df = pd.read_excel(file_name, engine='openpyxl')
                    except Exception:
                        df = pd.DataFrame()

                total_sales = df["الإجمالي"].sum() if "الإجمالي" in df.columns else 0
                total_qty = df["الكمية"].sum() if "الكمية" in df.columns else 0
                summary.append({"اسم العميل": cust, "إجمالي الكميات": total_qty, "إجمالي الحساب": total_sales})
                sheet = sanitize_sheet_name(cust)
                df.to_excel(writer, sheet_name=sheet, index=False)

            pd.DataFrame(summary).to_excel(writer, sheet_name="الملخص_العام", index=False)

        buffer.seek(0)
        return buffer
    finally:
        session.close()


# -------------------- واجهة المستخدم مع مصادقة بسيطة --------------------

init_db()

st.sidebar.header("الوصول")
if APP_PASSWORD:
    pwd = st.sidebar.text_input("أدخل كلمة المرور", type="password")
    if not pwd:
        st.sidebar.warning("هذه الواجهة محمية بكلمة مرور. الرجاء إدخالها للمتابعة.")
        st.stop()
    if pwd != APP_PASSWORD:
        st.sidebar.error("كلمة المرور غير صحيحة.")
        st.stop()
else:
    st.sidebar.info("لا توجد كلمة مرور معرفة. لو أردت حماية الواجهة ضع متغير البيئة APP_PASSWORD.")

st.sidebar.markdown("---")
st.sidebar.write("ملفات قاعدة البيانات: data.db (مخفي ضمن .gitignore).")

# الوضع العام
st.sidebar.header("الإعدادات")
mode = st.sidebar.radio("الوضع", ["عرض/تعديل", "إنشاء الملف الرئيسي / تنزيل"]) 

if mode == "عرض/تعديل":
    st.subheader("قائمة العملاء")
    customers = get_customers_db()
    if not customers:
        st.info("لم يتم العثور على عملاء بعد. يمكنك إضافة سجلات وسيتم إنشاء قاعدة البيانات تلقائيًا.")

    # حقل لاختيار أو إنشاء عميل جديد
    customer = st.selectbox("اختر اسم العميل:", options=[""] + customers, index=0)
    new_customer = st.text_input("أو اكتب اسم عميل جديد:")
    if new_customer and new_customer.strip():
        customer = new_customer.strip()

    st.markdown("---")
    st.subheader("إضافة صنف جديد")

    items = load_items()
    col1, col2 = st.columns([2, 1])
    with col1:
        item_type = st.selectbox("نوع الصنف:", options=[""] + items, index=0)
    with col2:
        new_item = st.text_input("أضف صنف جديد (إن رغبت)")

    qty = st.number_input("الكمية:", min_value=1, value=1, step=1)
    price = st.number_input("السعر للوحدة:", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    date_val = st.date_input("اختر التاريخ:", value=datetime.now())

    if st.button("حفظ الصنف"):
        if not customer:
            st.error("الرجاء اختيار اسم العميل أو كتابته في الحقل الأعلى.")
        else:
            if new_item and new_item.strip():
                try:
                    existing = load_items()
                    existing.append(new_item.strip())
                    with open(ITEMS_FILE, 'w', encoding='utf-8') as fh:
                        for it in sorted(set(existing)):
                            fh.write(it + '\n')
                    item_type = new_item.strip()
                    st.success("تم إضافة الصنف إلى القائمة.")
                except Exception as e:
                    st.error(f"خطأ عند تحديث قائمة الأصناف: {e}")
            if not item_type:
                st.error("الرجاء تحديد نوع الصنف.")
            else:
                date_str = date_val.strftime('%Y-%m-%d')
                ok = add_entry_db(customer, item_type, int(qty), float(price), date_str)
                if ok:
                    st.success("تم حفظ الصنف بنجاح.")
                    st.experimental_rerun()

    st.markdown("---")
    st.subheader("عرض وتحرير سجلات العميل")
    if customer:
        df = get_entries_by_customer(customer)
        if df.empty:
            st.info("لا توجد سجلات لهذا العميل بعد.")
        else:
            st.dataframe(df.drop(columns=[], errors='ignore'))
            # قائمة بالمعرفات للحذف
            ids = df['id'].tolist() if 'id' in df.columns else []
            if ids:
                chosen = st.selectbox("اختر ID للحذف:", options=[None] + ids)
                if st.button("حذف السجل المختار") and chosen:
                    if delete_entry_db(int(chosen)):
                        st.success("تم الحذف من قاعدة البيانات.")
                        st.experimental_rerun()
    else:
        st.info("اختر عميلًا من القائمة أو اكتب اسمًا جديدًا أعلاه.")

else:
    st.subheader("إنشاء الملف الرئيسي وتجميع أوراق العملاء")
    st.write("سيتم تجميع سجلات قاعدة البيانات وملفات Excel القديمة (إن وُجدت) في ملف Excel واحد يمكن تنزيله.")
    if st.button("إنشاء ملف رئيسي وتنزيله"):
        buffer = make_master_from_db_in_memory()
        if buffer is not None:
            st.download_button(label="تحميل الملف الرئيسي (Excel)", data=buffer.getvalue(), file_name=MASTER_FILE, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


st.sidebar.markdown("---")
st.sidebar.write("تشغيل محلي:\n1) تثبيت المتطلبات: `pip install -r requirements.txt`\n2) تشغيل: `streamlit run streamlit_app.py --server.headless=true`")
