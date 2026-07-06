import os
import glob
import re
from io import BytesIO
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np

# اسماء الملفات
ITEMS_FILE = "items_list.txt"
MASTER_FILE = "الملف_الرئيسي_المحاسبي.xlsx"

st.set_page_config(page_title="النظام المحاسبي - ويب", layout="centered")
st.title("النظام المحاسبي (نسخة ويب) — v2.2")

# -------------------- مساعدة --------------------

def sanitize_sheet_name(name: str) -> str:
    """تعقيم اسم ورقة Excel: يحذف الحروف الممنوعة ويقصر إلى 31 حرفًا."""
    if not isinstance(name, str):
        name = str(name)
    cleaned = re.sub(r'[:\\/*?\[\]]', '_', name)
    cleaned = cleaned.strip()
    if len(cleaned) > 31:
        cleaned = cleaned[:31]
    if not cleaned:
        cleaned = "sheet"
    return cleaned


def load_customers():
    files = glob.glob("*.xlsx")
    customers = []
    for f in files:
        if f.startswith("~$"):
            continue
        if os.path.basename(f) == MASTER_FILE:
            continue
        name = os.path.splitext(os.path.basename(f))[0]
        if name:
            customers.append(name)
    customers = sorted(set(customers))
    return customers


def load_items():
    items = []
    if os.path.exists(ITEMS_FILE):
        try:
            with open(ITEMS_FILE, 'r', encoding='utf-8') as fh:
                items = [line.strip() for line in fh if line.strip()]
        except Exception:
            items = []
    return sorted(set(items))


def save_item_to_file(customer: str, item_type: str, qty: int, price: float, date_str: str):
    file_name = f"{customer}.xlsx"
    total = qty * price
    new_row = {"نوع الصنف": item_type, "الكمية": qty, "السعر": price, "الإجمالي": total, "التاريخ": date_str}
    if os.path.exists(file_name):
        try:
            df = pd.read_excel(file_name, engine='openpyxl')
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        except Exception as e:
            st.error(f"خطأ: تعذر قراءة ملف العميل. تأكد أنه مغلق. \n{e}")
            return False
    else:
        df = pd.DataFrame([new_row])

    try:
        df.to_excel(file_name, index=False, engine='openpyxl')
        return True
    except Exception as e:
        st.error(f"خطأ عند الحفظ: {e}")
        return False


def delete_row_from_file(customer: str, idx):
    file_name = f"{customer}.xlsx"
    if not os.path.exists(file_name):
        st.error("ملف العميل غير موجود.")
        return False
    try:
        df = pd.read_excel(file_name, engine='openpyxl')
        if idx not in df.index:
            st.error("الفهرس المحدد غير صالح.")
            return False
        df = df.drop(index=idx).reset_index(drop=True)
        df.to_excel(file_name, index=False, engine='openpyxl')
        return True
    except Exception as e:
        st.error(f"خطأ عند الحذف: {e}")
        return False


def make_master_in_memory():
    files = glob.glob("*.xlsx")
    valid = [f for f in files if not f.startswith("~$") and os.path.basename(f) != MASTER_FILE]
    if not valid:
        st.warning("لا توجد ملفات عملاء لإنشاء الملف الرئيسي.")
        return None

    summary = []
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # جمع الملخص
        for f in valid:
            name = os.path.splitext(os.path.basename(f))[0]
            try:
                df = pd.read_excel(f, engine='openpyxl')
            except Exception:
                df = pd.DataFrame()
            total_sales = df["الإجمالي"].sum() if "الإجمالي" in df.columns else 0
            total_qty = df["الكمية"].sum() if "الكمية" in df.columns else 0
            summary.append({"اسم العميل": name, "إجمالي الكميات": total_qty, "إجمالي الحساب": total_sales})

        pd.DataFrame(summary).to_excel(writer, sheet_name="الملخص_العام", index=False)

        used = set()
        for f in valid:
            name = os.path.splitext(os.path.basename(f))[0]
            sheet = sanitize_sheet_name(name)
            # تحقيق التفرد
            if sheet in used:
                i = 1
                base = sheet[:28]
                while f"{base}_{i}" in used:
                    i += 1
                sheet = (base + f"_{i}")[:31]
            used.add(sheet)
            try:
                df = pd.read_excel(f, engine='openpyxl')
            except Exception:
                df = pd.DataFrame()
            df.to_excel(writer, sheet_name=sheet, index=False)

    buffer.seek(0)
    return buffer


# -------------------- واجهة المستخدم --------------------

st.sidebar.header("الإعدادات")
mode = st.sidebar.radio("الوضع", ["عرض/تعديل", "إنشاء الملف الرئيسي / تنزيل"])

if mode == "عرض/تعديل":
    st.subheader("قائمة العملاء")
    customers = load_customers()
    if not customers:
        st.info("لم يتم العثور على ملفات عملاء (.xlsx) في المستودع. يمكنك إضافة صفوف جديدة وسيتم إنشاء الملفات.")

    customer = st.selectbox("اختر اسم العميل:", options=[""] + customers, index=0)

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
        if new_item and new_item.strip():
            # حفظ في قائمة الأصناف
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
        if not customer:
            st.error("الرجاء اختيار اسم العميل (أو اكتب اسمًا في حقل الاسم أعلاه).")
        elif not item_type:
            st.error("الرجاء تحديد نوع الصنف.")
        else:
            date_str = date_val.strftime('%Y-%m-%d')
            ok = save_item_to_file(customer, item_type, int(qty), float(price), date_str)
            if ok:
                st.success("تم حفظ الصنف بنجاح.")
                st.experimental_rerun()

    st.markdown("---")
    st.subheader("عرض وتحرير جدول العميل")
    if customer:
        file_name = f"{customer}.xlsx"
        if os.path.exists(file_name):
            try:
                df = pd.read_excel(file_name, engine='openpyxl')
                if df.empty:
                    st.info("ملف العميل موجود لكن لا يحتوي على بيانات.")
                else:
                    st.dataframe(df)
                    st.write("اختر فهرس السطر لحذفه:")
                    idx = st.selectbox("فهرس السطر:", options=df.index.tolist())
                    if st.button("حذف السطر المحدد"):
                        if delete_row_from_file(customer, idx):
                            st.success("تم الحذف بنجاح.")
                            st.experimental_rerun()
            except Exception as e:
                st.error(f"تعذر قراءة ملف العميل: {e}")
        else:
            st.info("لا يوجد ملف لهذه العميل بعد. سيتم إنشاؤه عند حفظ أول صنف.")
    else:
        st.info("اختر عميلًا من القائمة أو أنشئ ملف عميل جديد بحفظ صنف له.")

else:
    st.subheader("إنشاء الملف الرئيسي وتجميع أوراق العملاء")
    st.write("سيتم تجميع ملفات العملاء الحالية في ملف Excel واحد يحتوي ورقة لكل عميل وورقة ملخص.")
    if st.button("إنشاء ملف رئيسي وتنزيله"):
        buffer = make_master_in_memory()
        if buffer is not None:
            st.download_button(label="تحميل الملف الرئيسي (Excel)", data=buffer.getvalue(), file_name=MASTER_FILE, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


st.sidebar.markdown("---")
st.sidebar.write("ملفات تُخزن محليًا داخل نفس مستودع المشروع. تأكد من نسخ البيانات احتياطيًا.")

st.sidebar.markdown("### تشغيل محلي\n1) تثبيت المتطلبات: `pip install -r requirements.txt`\n2) تشغيل: `streamlit run streamlit_app.py`")
