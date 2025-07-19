import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from datetime import datetime
import os

# برای Excel
try:
    import openpyxl
    from openpyxl.styles import PatternFill

    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# برای PDF
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.lib.units import cm

    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False


# ---------- تبدیل تاریخ شمسی به میلادی ساده (نمونه) ----------
def shamsi_to_gregorian(sh_date):
    """تبدیل ساده تاریخ - برای استفاده واقعی باید کتابخانه تبدیل تاریخ استفاده شود"""
    try:
        if not sh_date:
            return None
        parts = sh_date.split('/')
        if len(parts) != 3:
            return None
        y, m, d = map(int, parts)
        # برای نمونه، تاریخ شمسی را به میلادی تقریبی تبدیل می‌کنیم
        # در پیاده‌سازی واقعی باید از کتابخانه مناسب استفاده کرد
        gregorian_year = y + 621
        return datetime(gregorian_year, m, d)
    except:
        return None


# ---------- ذخیره و بارگذاری داده‌ها ----------
DATA_FILE = "projects_data.json"


def save_data(data):
    """ذخیره داده‌ها در فایل JSON"""
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در ذخیره داده‌ها: {str(e)}")


def load_data():
    """بارگذاری داده‌ها از فایل JSON"""
    try:
        if not os.path.exists(DATA_FILE):
            return []
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در بارگذاری داده‌ها: {str(e)}")
        return []


# ---------- تعیین وضعیت ----------
def determine_status(next_call_date, finished):
    """تعیین وضعیت پروژه بر اساس تاریخ تماس بعدی"""
    if finished:
        return ""
    if not next_call_date:
        return "انتظار"

    dt_today = datetime.now()
    dt_next_call = shamsi_to_gregorian(next_call_date)
    if dt_next_call is None:
        return "انتظار"
    if dt_next_call > dt_today:
        return "انتظار"
    else:
        return "در انتظار تماس مجدد"


# ---------- کلاس اصلی برنامه ----------
class ProjectManager:
    def __init__(self, root):
        self.root = root
        self.root.title("مدیریت پروژه‌ها")
        self.root.geometry("1400x800")

        # داده‌ها
        self.data = load_data()

        # متغیرها
        self.entries = {}
        self.finished_var = tk.BooleanVar()
        self.finished_status_var = tk.StringVar()

        # فیلترها
        self.filter_status_var = tk.StringVar()
        self.filter_name_var = tk.StringVar()
        self.filter_keyword_var = tk.StringVar()
        self.filter_date_from_var = tk.StringVar()
        self.filter_date_to_var = tk.StringVar()

        # مرتب‌سازی
        self.sort_by_var = tk.StringVar()
        self.sort_order_var = tk.StringVar()

        self.create_widgets()
        self.refresh_table()

    def create_widgets(self):
        """ایجاد عناصر واسط کاربری"""
        # فرم ورود داده
        self.create_form()

        # دکمه‌های عملیات
        self.create_buttons()

        # فیلتر و مرتب‌سازی
        self.create_filter_sort()

        # جدول نمایش داده‌ها
        self.create_table()

        # دکمه‌های خروجی
        self.create_export_buttons()

    def create_form(self):
        """ایجاد فرم ورود داده"""
        frame_form = tk.LabelFrame(self.root, text="ورود اطلاعات پروژه", font=("Arial", 10, "bold"))
        frame_form.pack(padx=10, pady=5, fill="x")

        # ردیف اول
        tk.Label(frame_form, text="نام مهندس:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.entries["name"] = tk.Entry(frame_form, width=20)
        self.entries["name"].grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        tk.Label(frame_form, text="آدرس:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.entries["address"] = tk.Entry(frame_form, width=30)
        self.entries["address"].grid(row=0, column=3, sticky="ew", padx=5, pady=2)

        # ردیف دوم
        tk.Label(frame_form, text="متراژ:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.entries["area"] = tk.Entry(frame_form, width=20)
        self.entries["area"].grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        tk.Label(frame_form, text="تعداد اتاق:").grid(row=1, column=2, sticky="w", padx=5, pady=2)
        self.entries["rooms"] = tk.Entry(frame_form, width=20)
        self.entries["rooms"].grid(row=1, column=3, sticky="ew", padx=5, pady=2)

        # ردیف سوم
        tk.Label(frame_form, text="تاریخ ویزیت (YYYY/MM/DD):").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.entries["visit_date"] = tk.Entry(frame_form, width=20)
        self.entries["visit_date"].grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        tk.Label(frame_form, text="تاریخ تماس بعدی (اختیاری):").grid(row=2, column=2, sticky="w", padx=5, pady=2)
        self.entries["next_call_date"] = tk.Entry(frame_form, width=20)
        self.entries["next_call_date"].grid(row=2, column=3, sticky="ew", padx=5, pady=2)

        # ردیف چهارم - توضیحات
        tk.Label(frame_form, text="توضیحات:").grid(row=3, column=0, sticky="nw", padx=5, pady=2)
        self.entries["description"] = tk.Text(frame_form, height=3, width=50)
        self.entries["description"].grid(row=3, column=1, columnspan=3, sticky="ew", padx=5, pady=2)

        # ردیف پنجم - وضعیت تمام شده
        finished_frame = tk.Frame(frame_form)
        finished_frame.grid(row=4, column=0, columnspan=4, sticky="ew", padx=5, pady=2)

        self.finished_check = tk.Checkbutton(finished_frame, text="تمام شده", variable=self.finished_var)
        self.finished_check.pack(side="left")

        tk.Label(finished_frame, text="وضعیت:").pack(side="left", padx=(20, 5))
        self.finished_status_dropdown = ttk.Combobox(finished_frame, textvariable=self.finished_status_var,
                                                     state="disabled", width=15)
        self.finished_status_dropdown['values'] = ("از دست رفته", "خرید")
        self.finished_status_dropdown.pack(side="left")

        tk.Label(finished_frame, text="تاریخ پایان:").pack(side="left", padx=(20, 5))
        self.entries["end_date"] = tk.Entry(finished_frame, width=15)
        self.entries["end_date"].pack(side="left")

        # اتصال رویداد تغییر وضعیت تمام شده
        self.finished_var.trace_add("write", self.on_finished_change)

        # تنظیم grid weights
        frame_form.columnconfigure(1, weight=1)
        frame_form.columnconfigure(3, weight=1)

    def create_buttons(self):
        """ایجاد دکمه‌های عملیات اصلی"""
        frame_buttons = tk.Frame(self.root)
        frame_buttons.pack(padx=10, pady=5, fill="x")

        tk.Button(frame_buttons, text="افزودن/ویرایش", command=self.add_or_update_entry,
                  bg="lightgreen").pack(side="left", padx=5)
        tk.Button(frame_buttons, text="حذف انتخاب شده", command=self.delete_selected,
                  bg="lightcoral").pack(side="left", padx=5)
        tk.Button(frame_buttons, text="پاک کردن فرم", command=self.clear_fields,
                  bg="lightblue").pack(side="left", padx=5)
        tk.Button(frame_buttons, text="بارگذاری در فرم", command=self.load_to_form,
                  bg="lightyellow").pack(side="left", padx=5)

    def create_filter_sort(self):
        """ایجاد بخش فیلتر و مرتب‌سازی"""
        frame_filter = tk.LabelFrame(self.root, text="فیلتر و مرتب‌سازی", font=("Arial", 10, "bold"))
        frame_filter.pack(padx=10, pady=5, fill="x")

        # ردیف اول فیلترها
        filter_row1 = tk.Frame(frame_filter)
        filter_row1.pack(fill="x", padx=5, pady=2)

        tk.Label(filter_row1, text="وضعیت:").pack(side="left")
        filter_status = ttk.Combobox(filter_row1, textvariable=self.filter_status_var, width=15)
        filter_status['values'] = ("همه", "از دست رفته", "خرید", "انتظار", "در انتظار تماس مجدد")
        filter_status.set("همه")
        filter_status.pack(side="left", padx=5)

        tk.Label(filter_row1, text="مهندس:").pack(side="left", padx=(20, 5))
        filter_name = tk.Entry(filter_row1, textvariable=self.filter_name_var, width=15)
        filter_name.pack(side="left", padx=5)

        tk.Label(filter_row1, text="کلیدواژه:").pack(side="left", padx=(20, 5))
        filter_keyword = tk.Entry(filter_row1, textvariable=self.filter_keyword_var, width=15)
        filter_keyword.pack(side="left", padx=5)

        # ردیف دوم فیلترها
        filter_row2 = tk.Frame(frame_filter)
        filter_row2.pack(fill="x", padx=5, pady=2)

        tk.Label(filter_row2, text="تاریخ از:").pack(side="left")
        filter_date_from = tk.Entry(filter_row2, textvariable=self.filter_date_from_var, width=12)
        filter_date_from.pack(side="left", padx=5)

        tk.Label(filter_row2, text="تا:").pack(side="left", padx=(5, 5))
        filter_date_to = tk.Entry(filter_row2, textvariable=self.filter_date_to_var, width=12)
        filter_date_to.pack(side="left", padx=5)

        # مرتب‌سازی
        tk.Label(filter_row2, text="مرتب‌سازی:").pack(side="left", padx=(20, 5))
        sort_by = ttk.Combobox(filter_row2, textvariable=self.sort_by_var, width=15)
        sort_by['values'] = ("تاریخ تماس بعدی", "تاریخ ویزیت", "تاریخ پایان", "نام مهندس")
        sort_by.set("تاریخ تماس بعدی")
        sort_by.pack(side="left", padx=5)

        sort_order = ttk.Combobox(filter_row2, textvariable=self.sort_order_var, width=10)
        sort_order['values'] = ("صعودی", "نزولی")
        sort_order.set("صعودی")
        sort_order.pack(side="left", padx=5)

        tk.Button(filter_row2, text="اعمال", command=self.apply_filter_sort,
                  bg="orange").pack(side="left", padx=10)
        tk.Button(filter_row2, text="پاک کردن فیلتر", command=self.clear_filters,
                  bg="lightgray").pack(side="left", padx=5)

    def create_table(self):
        """ایجاد جدول نمایش داده‌ها"""
        table_frame = tk.Frame(self.root)
        table_frame.pack(padx=10, pady=5, fill="both", expand=True)

        cols = ("نام مهندس", "آدرس", "متراژ", "تعداد اتاق", "تاریخ ویزیت",
                "تاریخ تماس بعدی", "وضعیت", "توضیحات", "تاریخ پایان")

        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=15)

        # تنظیم ستون‌ها
        column_widths = [120, 200, 80, 100, 120, 120, 150, 200, 120]
        for i, col in enumerate(cols):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths[i], anchor="center")

        # اسکرول بارها
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        # رنگ‌بندی
        self.tree.tag_configure("red", background="#f8d7da")  # از دست رفته
        self.tree.tag_configure("green", background="#d4edda")  # خرید
        self.tree.tag_configure("yellow", background="#fff3cd")  # انتظار
        self.tree.tag_configure("blue", background="#d1ecf1")  # در انتظار تماس

    def create_export_buttons(self):
        """ایجاد دکمه‌های خروجی"""
        export_frame = tk.Frame(self.root)
        export_frame.pack(padx=10, pady=5, fill="x")

        if EXCEL_AVAILABLE:
            tk.Button(export_frame, text="خروجی Excel", command=self.export_to_excel,
                      bg="lightgreen").pack(side="left", padx=5)
        else:
            tk.Button(export_frame, text="Excel غیرفعال (نیاز به openpyxl)",
                      state="disabled").pack(side="left", padx=5)

        if PDF_AVAILABLE:
            tk.Button(export_frame, text="خروجی PDF", command=self.export_to_pdf,
                      bg="lightcoral").pack(side="left", padx=5)
        else:
            tk.Button(export_frame, text="PDF غیرفعال (نیاز به reportlab)",
                      state="disabled").pack(side="left", padx=5)

    def on_finished_change(self, *args):
        """رویداد تغییر وضعیت تمام شده"""
        if self.finished_var.get():
            self.finished_status_dropdown.config(state="readonly")
        else:
            self.finished_status_var.set("")
            self.finished_status_dropdown.config(state="disabled")

    def clear_fields(self):
        """پاک کردن فیلدهای فرم"""
        for key, widget in self.entries.items():
            if isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
        self.finished_var.set(False)
        self.finished_status_var.set("")

    def clear_filters(self):
        """پاک کردن فیلترها"""
        self.filter_status_var.set("همه")
        self.filter_name_var.set("")
        self.filter_keyword_var.set("")
        self.filter_date_from_var.set("")
        self.filter_date_to_var.set("")
        self.refresh_table()

    def load_to_form(self):
        """بارگذاری رکورد انتخاب شده در فرم"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("اخطار", "لطفاً یک رکورد انتخاب کنید.")
            return

        item = self.tree.item(selected[0])
        values = item["values"]
        if not values:
            return

        # پیدا کردن رکورد در داده‌ها
        name = values[0]
        address = values[1]
        for rec in self.data:
            if rec.get("name") == name and rec.get("address") == address:
                self.entries["name"].delete(0, tk.END)
                self.entries["name"].insert(0, rec.get("name", ""))

                self.entries["address"].delete(0, tk.END)
                self.entries["address"].insert(0, rec.get("address", ""))

                self.entries["area"].delete(0, tk.END)
                self.entries["area"].insert(0, rec.get("area", ""))

                self.entries["rooms"].delete(0, tk.END)
                self.entries["rooms"].insert(0, rec.get("rooms", ""))

                self.entries["visit_date"].delete(0, tk.END)
                self.entries["visit_date"].insert(0, rec.get("visit_date", ""))

                self.entries["next_call_date"].delete(0, tk.END)
                self.entries["next_call_date"].insert(0, rec.get("next_call_date", ""))

                self.entries["description"].delete("1.0", tk.END)
                self.entries["description"].insert("1.0", rec.get("description", ""))

                self.entries["end_date"].delete(0, tk.END)
                self.entries["end_date"].insert(0, rec.get("end_date", ""))

                status = rec.get("status", "")
                if status in ("از دست رفته", "خرید"):
                    self.finished_var.set(True)
                    self.finished_status_var.set(status)
                else:
                    self.finished_var.set(False)
                    self.finished_status_var.set("")
                break

    def refresh_table(self, filtered_data=None):
        """بروزرسانی جدول"""
        for row in self.tree.get_children():
            self.tree.delete(row)

        display_data = filtered_data if filtered_data is not None else self.data

        for rec in display_data:
            vals = (
                rec.get("name", ""),
                rec.get("address", ""),
                rec.get("area", ""),
                rec.get("rooms", ""),
                rec.get("visit_date", ""),
                rec.get("next_call_date", ""),
                rec.get("status", ""),
                rec.get("description", "")[:50] + "..." if len(rec.get("description", "")) > 50 else rec.get(
                    "description", ""),
                rec.get("end_date", "")
            )

            # تعیین رنگ
            status = rec.get("status", "")
            tag = ""
            if status == "از دست رفته":
                tag = "red"
            elif status == "خرید":
                tag = "green"
            elif status == "انتظار":
                tag = "yellow"
            elif status == "در انتظار تماس مجدد":
                tag = "blue"

            self.tree.insert("", "end", values=vals, tags=(tag,))

    def add_or_update_entry(self):
        """افزودن یا ویرایش رکورد"""
        name = self.entries["name"].get().strip()
        address = self.entries["address"].get().strip()
        area = self.entries["area"].get().strip()
        rooms = self.entries["rooms"].get().strip()
        visit_date = self.entries["visit_date"].get().strip()
        next_call_date = self.entries["next_call_date"].get().strip()
        description = self.entries["description"].get("1.0", tk.END).strip()
        finished = self.finished_var.get()
        status = self.finished_status_var.get() if finished else ""
        end_date = self.entries["end_date"].get().strip() if finished else ""

        # اعتبارسنجی
        if not name or not visit_date:
            messagebox.showerror("خطا", "لطفاً حداقل نام مهندس و تاریخ ویزیت را وارد کنید.")
            return

        # بررسی فرمت تاریخ‌ها
        if shamsi_to_gregorian(visit_date) is None:
            messagebox.showerror("خطا", "فرمت تاریخ ویزیت صحیح نیست (YYYY/MM/DD).")
            return

        if next_call_date and shamsi_to_gregorian(next_call_date) is None:
            messagebox.showerror("خطا", "فرمت تاریخ تماس بعدی صحیح نیست (YYYY/MM/DD).")
            return

        if finished and end_date and shamsi_to_gregorian(end_date) is None:
            messagebox.showerror("خطا", "فرمت تاریخ پایان صحیح نیست (YYYY/MM/DD).")
            return

        # جستجو برای ویرایش
        found = False
        for rec in self.data:
            if rec.get("name") == name and rec.get("address") == address:
                rec.update({
                    "area": area,
                    "rooms": rooms,
                    "visit_date": visit_date,
                    "next_call_date": next_call_date,
                    "status": status if status else determine_status(next_call_date, finished),
                    "description": description,
                    "end_date": end_date
                })
                found = True
                break

        if not found:
            # افزودن رکورد جدید
            new_rec = {
                "name": name,
                "address": address,
                "area": area,
                "rooms": rooms,
                "visit_date": visit_date,
                "next_call_date": next_call_date,
                "status": status if status else determine_status(next_call_date, finished),
                "description": description,
                "end_date": end_date
            }
            self.data.append(new_rec)

        save_data(self.data)
        self.refresh_table()
        self.clear_fields()
        messagebox.showinfo("موفق", "رکورد با موفقیت ذخیره شد.")

    def delete_selected(self):
        """حذف رکورد انتخاب شده"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("اخطار", "لطفاً یک رکورد انتخاب کنید.")
            return

        if messagebox.askyesno("تایید حذف", "آیا مطمئن هستید که می‌خواهید این رکورد را حذف کنید؟"):
            item = self.tree.item(selected[0])
            values = item["values"]
            if not values:
                return

            name = values[0]
            address = values[1]
            self.data = [rec for rec in self.data if not (rec.get("name") == name and rec.get("address") == address)]
            save_data(self.data)
            self.refresh_table()
            messagebox.showinfo("موفق", "رکورد با موفقیت حذف شد.")

    def apply_filter_sort(self):
        """اعمال فیلتر و مرتب‌سازی"""
        filtered = []
        status_filter = self.filter_status_var.get()
        name_filter = self.filter_name_var.get().strip()
        keyword_filter = self.filter_keyword_var.get().strip()
        date_from_str = self.filter_date_from_var.get().strip()
        date_to_str = self.filter_date_to_var.get().strip()

        dt_from = shamsi_to_gregorian(date_from_str) if date_from_str else None
        dt_to = shamsi_to_gregorian(date_to_str) if date_to_str else None

        for rec in self.data:
            # فیلتر وضعیت
            if status_filter and status_filter != "همه" and rec.get("status") != status_filter:
                continue

            # فیلتر نام مهندس
            if name_filter and name_filter.lower() not in rec.get("name", "").lower():
                continue

            # فیلتر کلیدواژه در توضیحات
            if keyword_filter and keyword_filter.lower() not in rec.get("description", "").lower():
                continue

            # فیلتر بازه تاریخی
            next_call_date = rec.get("next_call_date", "")
            if next_call_date:
                dt_next_call = shamsi_to_gregorian(next_call_date)
                if dt_from and (dt_next_call is None or dt_next_call < dt_from):
                    continue
                if dt_to and (dt_next_call is None or dt_next_call > dt_to):
                    continue

            filtered.append(rec)

        # مرتب‌سازی
        sort_by = self.sort_by_var.get()
        sort_order = self.sort_order_var.get()
        reverse = sort_order == "نزولی"

        # کد جدید و کامل برای جایگزینی
        def get_sort_key(rec):
            if sort_by == "تاریخ تماس بعدی":
                dt = shamsi_to_gregorian(rec.get("next_call_date", ""))
                # برای مرتب‌سازی صحیح، اگر تاریخ نبود، یک تاریخ خیلی قدیمی برمی‌گردانیم
                return dt if dt is not None else datetime.min

            # این دو شرط elif به کد اضافه شده‌اند
            elif sort_by == "تاریخ ویزیت":
                dt = shamsi_to_gregorian(rec.get("visit_date", ""))
                return dt if dt is not None else datetime.min
            elif sort_by == "تاریخ پایان":
                dt = shamsi_to_gregorian(rec.get("end_date", ""))
                return dt if dt is not None else datetime.min

            elif sort_by == "نام مهندس":
                return rec.get("name", "")

            return ""

        filtered.sort(key=get_sort_key, reverse=reverse)
        self.refresh_table(filtered)

    def export_to_excel(self):
        """خروجی به فایل Excel"""
        if not EXCEL_AVAILABLE:
            messagebox.showerror("خطا", "کتابخانه openpyxl نصب نیست.")
            return

        if not self.data:
            messagebox.showinfo("اطلاع", "هیچ داده‌ای برای خروجی وجود ندارد.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not filepath:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "پروژه‌ها"

            # هدرها
            headers = ["نام مهندس", "آدرس", "متراژ", "تعداد اتاق", "تاریخ ویزیت",
                       "تاریخ تماس بعدی", "وضعیت", "توضیحات", "تاریخ پایان"]
            ws.append(headers)

            # رنگ‌های مختلف برای وضعیت‌ها
            fill_red = PatternFill(start_color="f8d7da", end_color="f8d7da", fill_type="solid")
            fill_green = PatternFill(start_color="d4edda", end_color="d4edda", fill_type="solid")
            fill_yellow = PatternFill(start_color="fff3cd", end_color="fff3cd", fill_type="solid")
            fill_blue = PatternFill(start_color="d1ecf1", end_color="d1ecf1", fill_type="solid")

            for rec in self.data:
                row_data = [
                    rec.get("name", ""),
                    rec.get("address", ""),
                    rec.get("area", ""),
                    rec.get("rooms", ""),
                    rec.get("visit_date", ""),
                    rec.get("next_call_date", ""),
                    rec.get("status", ""),
                    rec.get("description", ""),
                    rec.get("end_date", "")
                ]
                ws.append(row_data)

                # رنگ‌بندی بر اساس وضعیت
                fill = None
                status = rec.get("status", "")
                if status == "از دست رفته":
                    fill = fill_red
                elif status == "خرید":
                    fill = fill_green
                elif status == "انتظار":
                    fill = fill_yellow
                elif status == "در انتظار تماس مجدد":
                    fill = fill_blue

                if fill:
                    for cell in ws[ws.max_row]:
                        cell.fill = fill

            # تنظیم عرض ستون‌ها
            column_widths = [15, 25, 10, 12, 15, 15, 20, 30, 15]
            for i, width in enumerate(column_widths, 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width

            wb.save(filepath)
            messagebox.showinfo("موفق", f"فایل Excel با موفقیت در {filepath} ذخیره شد.")

        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ایجاد فایل Excel: {str(e)}")

    def export_to_pdf(self):
        """خروجی به فایل PDF"""
        if not PDF_AVAILABLE:
            messagebox.showerror("خطا", "کتابخانه reportlab نصب نیست.")
            return

        if not self.data:
            messagebox.showinfo("اطلاع", "هیچ داده‌ای برای خروجی وجود ندارد.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not filepath:
            return

        try:
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4
            margin = 2 * cm
            y = height - margin

            # عنوان
            c.setFont("Helvetica-Bold", 16)
            c.drawRightString(width - margin, y, "گزارش مدیریت پروژه‌ها")
            y -= 1.5 * cm

            c.setFont("Helvetica", 8)
            c.drawRightString(width - margin, y, f"تاریخ تولید: {datetime.now().strftime('%Y/%m/%d %H:%M')}")
            y -= 2 * cm

            # هدر جدول
            c.setFont("Helvetica-Bold", 9)
            headers = ["نام مهندس", "آدرس", "متراژ", "اتاق", "تاریخ ویزیت", "تماس بعدی", "وضعیت"]
            col_widths = [2.5 * cm, 4 * cm, 1.5 * cm, 1 * cm, 2.5 * cm, 2.5 * cm, 3 * cm]

            x = margin
            for i, header in enumerate(headers):
                c.drawString(x, y, header)
                x += col_widths[i]
            y -= 0.5 * cm

            # خط جداکننده
            c.line(margin, y, width - margin, y)
            y -= 0.5 * cm

            # داده‌ها
            c.setFont("Helvetica", 8)
            row_height = 0.8 * cm

            for rec in self.data:
                # بررسی فضای باقی‌مانده در صفحه
                if y < margin + 2 * cm:
                    c.showPage()
                    y = height - margin

                # رنگ پس‌زمینه بر اساس وضعیت
                status = rec.get("status", "")
                if status == "از دست رفته":
                    c.setFillColor(colors.HexColor("#f8d7da"))
                elif status == "خرید":
                    c.setFillColor(colors.HexColor("#d4edda"))
                elif status == "انتظار":
                    c.setFillColor(colors.HexColor("#fff3cd"))
                elif status == "در انتظار تماس مجدد":
                    c.setFillColor(colors.HexColor("#d1ecf1"))
                else:
                    c.setFillColor(colors.white)

                # رسم پس‌زمینه
                c.rect(margin, y - 0.2 * cm, sum(col_widths), row_height, fill=1, stroke=0)
                c.setFillColor(colors.black)

                # رسم متن
                values = [
                    rec.get("name", "")[:15],
                    rec.get("address", "")[:25],
                    rec.get("area", ""),
                    rec.get("rooms", ""),
                    rec.get("visit_date", ""),
                    rec.get("next_call_date", ""),
                    rec.get("status", "")
                ]

                x = margin
                for i, value in enumerate(values):
                    c.drawString(x, y, str(value))
                    x += col_widths[i]

                y -= row_height

            c.save()
            messagebox.showinfo("موفق", f"فایل PDF با موفقیت در {filepath} ذخیره شد.")

        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ایجاد فایل PDF: {str(e)}")


def main():
    """تابع اصلی برنامه"""
    root = tk.Tk()
    app = ProjectManager(root)

    # تنظیمات اضافی
    root.protocol("WM_DELETE_WINDOW", lambda: (save_data(app.data), root.destroy()))

    try:
        root.mainloop()
    except KeyboardInterrupt:
        save_data(app.data)
        root.destroy()


if __name__ == "__main__":
    main()