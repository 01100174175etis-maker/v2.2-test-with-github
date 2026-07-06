import os
import glob
import re
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
import pandas as pd
import numpy as np

# استيراد DateEntry إن كانت متاحة، وإلا نوفّر بديلًا بسيطًا
try:
    from tkcalendar import DateEntry  # type: ignore
except Exception:
    DateEntry = None

ITEMS_FILE = "items_list.txt"


def debug_print(*args, **kwargs):
    # إرسال أخطاء وdebug إلى stderr ليظهر عند التشغيل من الطرفية
    print(*args, file=sys.stderr, **kwargs)


# بديل بسيط لمربع التاريخ إذا لم تتوفر tkcalendar
class FallbackDateEntry(ttk.Entry):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        # افتراض قيمة اليوم كقيمة أولية
        self.insert(0, datetime.now().strftime("%Y-%m-%d"))

    def get_date(self):
        txt = self.get().strip()
        try:
            # لقبول صيغ YYYY-MM-DD
            dt = datetime.strptime(txt, "%Y-%m-%d")
            return dt
        except Exception:
            # ضبط إلى الآن إذا فشل التحويل
            return datetime.now()


def sanitize_sheet_name(name: str) -> str:
    # تمنع الحروف التالية: : \ / ? * [ ]
    if not isinstance(name, str) or not name:
        return "sheet"
    cleaned = re.sub(r'[:\\/*?\[\]]', '_', name)
    cleaned = cleaned.strip()
    # الحد الأقصى لأسماء أوراق الإكسل 31 حرفًا
    if len(cleaned) > 31:
        cleaned = cleaned[:31]
    if not cleaned:
        return "sheet"
    return cleaned


# 1. دالة جلب قائمة العملاء بأمان من أسماء الملفات فقط
def load_customers():
    try:
        files = glob.glob("*.xlsx")
        customer_set = set()
        for file in files:
            if not file.startswith("~$") and file != "الملف_الرئيسي_المحاسبي.xlsx":
                name = os.path.splitext(file)[0]
                if name:
                    customer_set.add(name)
        customer_list = sorted(customer_set, key=lambda s: s)
        combo_customer['values'] = customer_list
    except Exception as e:
        debug_print("load_customers error:", e)


# 2. دالة جلب وتحديث قائمة الأصناف من ملف نصي خارجي
def load_items(new_item=None):
    items = set()
    try:
        if os.path.exists(ITEMS_FILE):
            with open(ITEMS_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    v = line.strip()
                    if v:
                        items.add(v)
    except Exception as e:
        debug_print("load_items read error:", e)

    if new_item and new_item.strip():
        items.add(new_item.strip())
        try:
            with open(ITEMS_FILE, "w", encoding="utf-8") as f:
                for item in sorted(items):
                    f.write(f"{item}\n")
        except Exception as e:
            debug_print("load_items write error:", e)

    combo_type['values'] = sorted(items)


# 3. دالة قراءة وعرض جدول العميل عند اختياره
def show_customer_table(event=None):
    customer = combo_customer.get().strip()
    file_name = f"{customer}.xlsx"

    for item in tree.get_children():
        tree.delete(item)

    if not customer or not os.path.exists(file_name):
        return

    try:
        df = pd.read_excel(file_name, engine='openpyxl')
        # نعوّض القيم الفارغة لتجنب ظهور NaN
        df = df.fillna("")
        for _, row in df.iterrows():
            # استخدام get مع fallback لتجنب KeyError
            date_val = row.get("التاريخ", "")
            total_val = row.get("الإجمالي", "")
            price_val = row.get("السعر", "")
            qty_val = row.get("الكمية", "")
            type_val = row.get("نوع الصنف", "")
            # تحويل لقيم نصية للعرض
            tree.insert("", tk.END, values=(str(date_val), str(total_val), str(price_val), str(qty_val), str(type_val)))
    except Exception as e:
        messagebox.showerror("خطأ", f"تعذر فتح ملف العميل: {e}")
        debug_print("show_customer_table error:", e)


# 4. دالة حفظ البيانات الأساسية
def save_data():
    customer = combo_customer.get().strip()
    item_type = combo_type.get().strip()
    qty_str = ent_qty.get().strip()
    price_str = ent_price.get().strip()

    # الحصول على التاريخ
    try:
        if DateEntry is not None and isinstance(cal_date, DateEntry):
            date_dt = cal_date.get_date()
        else:
            date_dt = cal_date.get_date()
        date_str = date_dt.strftime('%Y-%m-%d')
    except Exception:
        date_str = datetime.now().strftime('%Y-%m-%d')

    if not customer or not item_type or not qty_str or not price_str:
        messagebox.showwarning("تنبيه", "الرجاء كتابة اسم العميل، الصنف، الكمية والسعر!")
        return

    try:
        qty = int(qty_str)
        price = float(price_str)
    except ValueError:
        messagebox.showerror("خطأ", "يجب إدخال أرقام صحيحة للكمية والسعر.")
        return

    total = qty * price
    file_name = f"{customer}.xlsx"

    new_row = {
        "نوع الصنف": item_type,
        "الكمية": qty,
        "السعر": price,
        "الإجمالي": total,
        "التاريخ": date_str
    }

    try:
        if os.path.exists(file_name):
            try:
                df = pd.read_excel(file_name, engine='openpyxl')
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            except Exception as e:
                messagebox.showerror("خطأ", "تأكد من إغلاق ملف إكسل العميل أولاً!")
                debug_print("save_data read existing file error:", e)
                return
        else:
            df = pd.DataFrame([new_row])

        # حفظ الملف
        df.to_excel(file_name, index=False, engine='openpyxl')
        messagebox.showinfo("نجاح", "تم حفظ الصنف بنجاح.")

        # تحديث القوائم والواجهة
        load_items(item_type)
        load_customers()
        combo_customer.set(customer)
        show_customer_table()

        combo_type.set('')
        ent_qty.delete(0, tk.END)
        ent_price.delete(0, tk.END)
    except Exception as e:
        messagebox.showerror("خطأ في الحفظ", f"لم يتم الحفظ: {e}")
        debug_print("save_data write error:", e)


# 5. دالة حذف السطر المحدد من الجدول ومن ملف الإكسل
def delete_selected_item():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("تنبيه", "الرجاء تحديد السطر المراد حذفه من الجدول أولاً!")
        return

    customer = combo_customer.get().strip()
    file_name = f"{customer}.xlsx"

    if not customer or not os.path.exists(file_name):
        return

    confirm = messagebox.askyesno("تأكيد الحذف", "هل أنت متأكد من رغبتك في حذف هذا الصنف نهائياً؟")
    if not confirm:
        return

    try:
        # selection ترجع قائمة، نتعامل مع العنصر الأول المحدد
        sel = selected_item[0]
        item_values = tree.item(sel)['values']
        if not item_values or len(item_values) < 5:
            messagebox.showerror("خطأ", "قيمة السطر المحدد غير صحيحة.")
            return

        target_date = str(item_values[0])
        target_total = float(item_values[1]) if item_values[1] != "" else None
        target_price = float(item_values[2]) if item_values[2] != "" else None
        target_qty = int(float(item_values[3])) if item_values[3] != "" else None
        target_type = str(item_values[4])

        df = pd.read_excel(file_name, engine='openpyxl')

        # معالجة قيم التاريخ في الملف إلى صيغة نصية موحدة
        if "التاريخ" in df.columns:
            df["__التاريخ_str__"] = pd.to_datetime(df["التاريخ"], errors='coerce').dt.strftime('%Y-%m-%d').fillna(df["التاريخ"].astype(str))
        else:
            df["__التاريخ_str__"] = ""

        # بناء شرط متدرج مع مراعاة المقارنة العشرية للـ price و total
        condition = pd.Series([True] * len(df))
        try:
            condition &= (df.get("نوع الصنف", "").astype(str) == target_type)
        except Exception:
            condition &= False

        if target_qty is not None and "الكمية" in df.columns:
            try:
                condition &= (df["الكمية"].astype(float) == float(target_qty))
            except Exception:
                condition &= False

        if target_price is not None and "السعر" in df.columns:
            try:
                # مقارنة تقريبية للأرقام العشرية
                condition &= np.isclose(df["السعر"].astype(float), float(target_price))
            except Exception:
                condition &= False

        if target_date:
            try:
                condition &= (df["__التاريخ_str__"].astype(str) == target_date)
            except Exception:
                condition &= False

        # نتحقق إن كان هناك أي صفوف مطابقة
        if condition.any():
            idx_to_drop = df[condition].index
            df = df.drop(idx_to_drop).reset_index(drop=True)
            # ازالة العمود المؤقت إن وجد
            if "__التاريخ_str__" in df.columns:
                try:
                    df = df.drop(columns="__التاريخ_str__")
                except Exception:
                    pass
            df.to_excel(file_name, index=False, engine='openpyxl')
            messagebox.showinfo("نجاح", "تم حذف الصنف وتحديث الملف بنجاح.")
            load_customers()
            show_customer_table()
        else:
            messagebox.showerror("خطأ", "تعذر العثور على السطر المطابق في ملف الـ Excel.")
    except Exception as e:
        messagebox.showerror("خطأ في الحذف", f"فشل حذف الصنف: {e}")
        debug_print("delete_selected_item error:", e)


# 6. دالة إنشاء الملف الرئيسي بصفحة لكل عميل + صفحة الملخص العام
def make_master_file():
    files = glob.glob("*.xlsx")
    master_file = "الملف_الرئيسي_المحاسبي.xlsx"
    valid_files = [f for f in files if not f.startswith("~$") and f != master_file]

    if not valid_files:
        messagebox.showwarning("تنبيه", "لا توجد ملفات عملاء حالياً!")
        return

    summary = []
    try:
        with pd.ExcelWriter(master_file, engine='openpyxl') as writer:
            for file in valid_files:
                name = os.path.splitext(file)[0]
                try:
                    df = pd.read_excel(file, engine='openpyxl')
                except Exception:
                    df = pd.DataFrame()

                total_sales = df["الإجمالي"].sum() if "الإجمالي" in df.columns else 0
                total_qty = df["الكمية"].sum() if "الكمية" in df.columns else 0

                summary.append({
                    "اسم العميل": name,
                    "إجمالي الكميات": total_qty,
                    "إجمالي الحساب": total_sales
                })

            pd.DataFrame(summary).to_excel(writer, sheet_name="الملخص_العام", index=False)

            # كتابة كل ورقة عميل مع تعقيم الاسم
            for file in valid_files:
                name = os.path.splitext(file)[0]
                sheet_name = sanitize_sheet_name(name)
                try:
                    df = pd.read_excel(file, engine='openpyxl')
                except Exception:
                    df = pd.DataFrame()
                try:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                except Exception as e:
                    debug_print(f"make_master_file: can't write sheet {sheet_name}:", e)

        messagebox.showinfo("نجاح", f"تم إنشاء الملف الرئيسي الشامل بنجاح باسم:\n{master_file}")
    except Exception as e:
        messagebox.showerror("خطأ", f"تأكد من إغلاق الملف الرئيسي أولاً: {e}")
        debug_print("make_master_file error:", e)


# ----------------- بناء الواجهة الرسومية -----------------
root = tk.Tk()
root.title("النظام المحاسبي - إصدار محسن")
root.geometry("460x700")

# إطار العميل
frame_c = tk.LabelFrame(root, text=" العميل ", padx=10, pady=5)
frame_c.pack(fill="x", padx=10, pady=5)

combo_customer = ttk.Combobox(frame_c, justify="right", width=30)
combo_customer.pack(side="right", padx=5)
combo_customer.bind("<<ComboboxSelected>>", show_customer_table)

lbl_c = tk.Label(frame_c, text="اسم العميل:")
lbl_c.pack(side="right")

btn_ref = tk.Button(frame_c, text="عرض الجدول", command=show_customer_table)
btn_ref.pack(side="left")

# إطار تفاصيل الصنف
frame_i = tk.LabelFrame(root, text=" تفاصيل الصنف ", padx=10, pady=5)
frame_i.pack(fill="x", padx=10, pady=5)

combo_type = ttk.Combobox(frame_i, justify="right", width=25)
combo_type.grid(row=0, column=0, pady=5, padx=5)
tk.Label(frame_i, text="نوع الصنف:").grid(row=0, column=1, sticky="e", padx=5)

ent_qty = tk.Entry(frame_i, justify="right", width=28)
ent_qty.grid(row=1, column=0, pady=5, padx=5)
tk.Label(frame_i, text="الكمية:").grid(row=1, column=1, sticky="e", padx=5)

ent_price = tk.Entry(frame_i, justify="right", width=28)
ent_price.grid(row=2, column=0, pady=5, padx=5)
tk.Label(frame_i, text="السعر لِلْوحدة:").grid(row=2, column=1, sticky="e", padx=5)

# مربع التقويم: استخدم DateEntry إن وُجد، وإلا البديل
if DateEntry is not None:
    cal_date = DateEntry(frame_i, width=25, background='darkgreen',
                         foreground='white', borderwidth=2,
                         year=datetime.now().year,
                         date_pattern='yyyy-mm-dd', justify="center")
else:
    cal_date = FallbackDateEntry(frame_i, width=25, justify="center")

cal_date.grid(row=3, column=0, pady=5, padx=5)
tk.Label(frame_i, text="اختر التاريخ:").grid(row=3, column=1, sticky="e", padx=5)

# الأزرار الرئيسية
btn_save = tk.Button(root, text="تسجيل و حفظ", bg="#5a376e", fg="white", width=40, command=save_data)
btn_save.pack(pady=4)

btn_master = tk.Button(root, text="تحديث وتصدير الملف الرئيسي الشامل", bg="#0065a3", fg="white", width=40, command=make_master_file)
btn_master.pack(pady=4)

# جدول العرض بالواجهة
frame_t = tk.LabelFrame(root, text=" جدول حسابات العميل الحالي ")
frame_t.pack(fill="both", expand=True, padx=10, pady=5)

columns = ("التاريخ", "الإجمالي", "السعر", "الكمية", "نوع الصنف")
tree = ttk.Treeview(frame_t, columns=columns, show="headings", height=8)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=110)
tree.pack(fill="both", expand=True, padx=5, pady=5)

# زر الحذف
btn_delete = tk.Button(root, text="حذف الصنف المحدد من ملف العميل", bg="#e74c3c", fg="white", width=40, command=delete_selected_item)
btn_delete.pack(pady=8)

# استدعاء الإعدادات الأولية
load_customers()
load_items()

root.mainloop()
