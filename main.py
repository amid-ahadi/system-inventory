import tkinter as tk
from tkinter import ttk, messagebox, Toplevel, Entry, BooleanVar
import sqlite3
import platform
import socket
import uuid
import psutil
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment

# ————————————————————————————————————————————————————————————————
# تنظیمات دیتابیس
# ————————————————————————————————————————————————————————————————

DB_NAME = "system_inventory_basic.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS devices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            location TEXT,
            hostname TEXT,
            case_model TEXT,
            monitor_model TEXT,
            keyboard_model TEXT,
            mouse_model TEXT,
            printer_model TEXT,
            internal_phone TEXT,
            delivery_date TEXT,
            recipient_name TEXT,
            ipv4_address TEXT,
            ipv6_address TEXT,
            mac_address TEXT,
            switch_connected TEXT,
            motherboard_model TEXT,
            has_dvd BOOLEAN,
            disk_type_capacity TEXT,
            ram_capacity_gb REAL,
            graphics_card TEXT,
            os_info TEXT,
            special_software TEXT,
            case_appearance TEXT,
            monitor_appearance TEXT,
            keyboard_appearance TEXT,
            printer_appearance TEXT,
            cpu_model TEXT,
            cpu_count INTEGER,
            ram_total_gb REAL,
            gpu_model TEXT,
            disk_total_gb REAL,
            purchase_date TEXT,
            notes TEXT,
            collection_date TEXT
        )
    ''')
    conn.commit()
    conn.close()

# ————————————————————————————————————————————————————————————————
# تابع جمع‌آوری اطلاعات سیستم
# ————————————————————————————————————————————————————————————————

def get_system_info():
    info = {}
    # نام کامپیوتر
    info['hostname'] = socket.gethostname()
    # نام سیستم عامل
    info['os_info'] = platform.platform()
    # نوع CPU
    info['cpu_model'] = platform.processor() or "نامشخص"
    # تعداد هسته
    info['cpu_count'] = psutil.cpu_count(logical=True)
    # مقدار RAM (به گیگابایت)
    ram = psutil.virtual_memory()
    info['ram_total_gb'] = round(ram.total / (1024 ** 3), 2)
    # کل فضای دیسک (تمام درایورها)
    total_disk = 0
    for partition in psutil.disk_partitions():
        try:
            usage = psutil.disk_usage(partition.mountpoint)
            total_disk += usage.total
        except PermissionError:
            continue
    info['disk_total_gb'] = round(total_disk / (1024 ** 3), 2)
    # آدرس IP
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        ip = "127.0.0.1"
    info['ipv4_address'] = ip
    # آدرس MAC
    info['mac_address'] = ':'.join(f'{(uuid.getnode() >> e) & 0xff:02x}' for e in range(0, 2 * 6, 2))
    # IPv6 (اولین آدرس IPv6 غیر لوکال)
    info['ipv6_address'] = "نامشخص"
    try:
        for iface, addrs in psutil.net_if_addrs().items():
            for addr in addrs:
                if addr.family == socket.AF_INET6 and not addr.address.startswith("fe80") and not addr.address.startswith("::1"):
                    info['ipv6_address'] = addr.address.split('%')[0]  # حذف scope
                    break
            if info['ipv6_address'] != "نامشخص":
                break
    except Exception:
        pass
    # GPU (عمومی)
    try:
        if hasattr(psutil, 'sensors_temperatures'):
            temps = psutil.sensors_temperatures()
            gpus = [name for name in temps.keys() if 'gpu' in name.lower() or 'amdgpu' in name.lower()]
            info['gpu_model'] = gpus[0].title() if gpus else "نامشخص (سنسور دما یافت نشد)"
        else:
            info['gpu_model'] = "نامشخص (سنسور دما پشتیبانی نمی‌شود)"
    except Exception:
        info['gpu_model'] = "نامشخص (خطا در خواندن سنسور)"
    # زمان جمع‌آوری
    info['collection_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return info

# ————————————————————————————————————————————————————————————————
# ایجاد GUI
# ————————————————————————————————————————————————————————————————

class SystemInventoryBasicApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ثبت اطلاعات سیستم")
        self.root.geometry("750x660")
        self.root.resizable(False, False)

        # متغیرهای ورودی کاربر
        self.name_var = tk.StringVar()
        self.location_var = tk.StringVar()
        self.case_model_var = tk.StringVar()
        self.monitor_model_var = tk.StringVar()
        self.keyboard_model_var = tk.StringVar()
        self.mouse_model_var = tk.StringVar()
        self.printer_model_var = tk.StringVar()
        self.internal_phone_var = tk.StringVar()
        self.delivery_date_var = tk.StringVar()
        self.recipient_name_var = tk.StringVar()
        self.switch_connected_var = tk.StringVar()
        self.motherboard_model_var = tk.StringVar()
        self.has_dvd_var = BooleanVar(value=False)
        self.disk_type_capacity_var = tk.StringVar()
        self.ram_capacity_gb_var = tk.StringVar()
        self.graphics_card_var = tk.StringVar()
        self.special_software_var = tk.StringVar()
        self.case_appearance_var = tk.StringVar()
        self.monitor_appearance_var = tk.StringVar()
        self.keyboard_appearance_var = tk.StringVar()
        self.printer_appearance_var = tk.StringVar()
        self.purchase_date_var = tk.StringVar()
        self.notes_var = tk.StringVar()

        # اطلاعات خودکار
        self.system_info = get_system_info()

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")

        # عنوان
        title_label = ttk.Label(main_frame, text="ثبت اطلاعات سیستم", font=("Tahoma", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 20))

        # نام دستگاه و محل اسکان
        ttk.Label(main_frame, text="نام كاربر:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.name_var, width=20).grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="محل اسکان:").grid(row=1, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.location_var, width=20).grid(row=1, column=3, sticky="w", padx=5, pady=5)

        # مدل‌های دستگاه
        ttk.Label(main_frame, text="مدل کیس:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.case_model_var, width=20).grid(row=2, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="مدل مانیتور:").grid(row=2, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.monitor_model_var, width=20).grid(row=2, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="مدل کیبورد:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.keyboard_model_var, width=20).grid(row=3, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="مدل ماوس:").grid(row=3, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.mouse_model_var, width=20).grid(row=3, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="مدل چاپگر:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.printer_model_var, width=20).grid(row=4, column=1, sticky="w", padx=5, pady=5)

        # شماره داخلی
        ttk.Label(main_frame, text="شماره داخلی:").grid(row=4, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.internal_phone_var, width=20).grid(row=4, column=3, sticky="w", padx=5, pady=5)

        # تاریخ تحویل و تحویل گیرنده
        ttk.Label(main_frame, text="تاریخ تحویل:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.delivery_date_var, width=20).grid(row=5, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="تحویل گیرنده:").grid(row=5, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.recipient_name_var, width=20).grid(row=5, column=3, sticky="w", padx=5, pady=5)

        # سوئیچ متصل و مادربرد
        ttk.Label(main_frame, text="سوئیچ متصل:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.switch_connected_var, width=20).grid(row=6, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="مدل مادربرد:").grid(row=6, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.motherboard_model_var, width=20).grid(row=6, column=3, sticky="w", padx=5, pady=5)

        # DVD دارد یا ندارد
        ttk.Label(main_frame, text="DVD دارد؟:").grid(row=7, column=0, sticky="e", padx=5, pady=5)
        ttk.Checkbutton(main_frame, variable=self.has_dvd_var).grid(row=7, column=1, sticky="w", padx=5, pady=5)

        # هارد دیسک و رم
        ttk.Label(main_frame, text="هارد/SSD (مثال: ssd 128):").grid(row=7, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.disk_type_capacity_var, width=20).grid(row=7, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="رم (GB):").grid(row=8, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.ram_capacity_gb_var, width=20).grid(row=8, column=1, sticky="w", padx=5, pady=5)

        # کارت گرافیک و نرم‌افزارهای خاص
        ttk.Label(main_frame, text="کارت گرافیک:").grid(row=8, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.graphics_card_var, width=20).grid(row=8, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="نرم‌افزارهای خاص:").grid(row=9, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.special_software_var, width=20).grid(row=9, column=1, sticky="w", padx=5, pady=5)

        # ظاهرها
        ttk.Label(main_frame, text="ظاهر کیس:").grid(row=9, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.case_appearance_var, width=20).grid(row=9, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="ظاهر مانیتور:").grid(row=10, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.monitor_appearance_var, width=20).grid(row=10, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="ظاهر کیبورد:").grid(row=10, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.keyboard_appearance_var, width=20).grid(row=10, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="ظاهر چاپگر:").grid(row=11, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.printer_appearance_var, width=20).grid(row=11, column=1, sticky="w", padx=5, pady=5)

        # تاریخ خرید
        ttk.Label(main_frame, text="تاریخ خرید:").grid(row=11, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(main_frame, textvariable=self.purchase_date_var, width=20).grid(row=11, column=3, sticky="w", padx=5, pady=5)

        # یادداشت‌ها
        ttk.Label(main_frame, text="یادداشت‌ها:").grid(row=12, column=0, sticky="ne", padx=5, pady=5)
        notes_entry = tk.Text(main_frame, height=4, width=50)
        notes_entry.grid(row=12, column=1, columnspan=3, sticky="w", padx=5, pady=5)
        notes_entry.bind("<KeyRelease>", lambda e: self.notes_var.set(notes_entry.get("1.0", "end-1c")))

        # نمایش اطلاعات خودکار
        ttk.Label(main_frame, text="نام کامپیوتر:").grid(row=13, column=0, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=self.system_info['hostname']).grid(row=13, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="IPv4:").grid(row=13, column=2, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=self.system_info['ipv4_address']).grid(row=13, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="IPv6:").grid(row=14, column=0, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=self.system_info['ipv6_address']).grid(row=14, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="MAC:").grid(row=14, column=2, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=self.system_info['mac_address']).grid(row=14, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="سیستم عامل:").grid(row=15, column=0, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=self.system_info['os_info']).grid(row=15, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="پردازنده:").grid(row=15, column=2, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=self.system_info['cpu_model']).grid(row=15, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="تعداد هسته:").grid(row=16, column=0, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=str(self.system_info['cpu_count'])).grid(row=16, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="RAM (GB) - خودکار:").grid(row=16, column=2, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=str(self.system_info['ram_total_gb'])).grid(row=16, column=3, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="کل فضای دیسک (GB):").grid(row=17, column=0, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=str(self.system_info['disk_total_gb'])).grid(row=17, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(main_frame, text="GPU:").grid(row=17, column=2, sticky="e", padx=5, pady=5)
        ttk.Label(main_frame, text=self.system_info['gpu_model']).grid(row=17, column=3, sticky="w", padx=5, pady=5)

        # دکمه ثبت
        submit_btn = ttk.Button(main_frame, text="ثبت اطلاعات", command=self.submit_data)
        submit_btn.grid(row=18, column=0, columnspan=2, pady=10)

        # دکمه خروجی اکسل
        export_btn = ttk.Button(main_frame, text="خروجی اکسل", command=self.open_password_window)
        export_btn.grid(row=18, column=2, columnspan=2, pady=10)

    def submit_data(self):
        if not self.name_var.get():
            messagebox.showwarning("هشدار", "لطفاً نام دستگاه را وارد کنید.")
            return
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()

            cursor.execute('''
                INSERT INTO devices (
                    name, location, hostname, case_model, monitor_model,
                    keyboard_model, mouse_model, printer_model, internal_phone,
                    delivery_date, recipient_name, ipv4_address, ipv6_address,
                    mac_address, switch_connected, motherboard_model, has_dvd,
                    disk_type_capacity, ram_capacity_gb, graphics_card,
                    os_info, special_software, case_appearance, monitor_appearance,
                    keyboard_appearance, printer_appearance, cpu_model, cpu_count,
                    ram_total_gb, gpu_model, disk_total_gb, purchase_date,
                    notes, collection_date
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                self.name_var.get(),
                self.location_var.get(),
                self.system_info['hostname'],
                self.case_model_var.get(),
                self.monitor_model_var.get(),
                self.keyboard_model_var.get(),
                self.mouse_model_var.get(),
                self.printer_model_var.get(),
                self.internal_phone_var.get(),
                self.delivery_date_var.get(),
                self.recipient_name_var.get(),
                self.system_info['ipv4_address'],
                self.system_info['ipv6_address'],
                self.system_info['mac_address'],
                self.switch_connected_var.get(),
                self.motherboard_model_var.get(),
                self.has_dvd_var.get(),
                self.disk_type_capacity_var.get(),
                float(self.ram_capacity_gb_var.get()) if self.ram_capacity_gb_var.get() else 0.0,
                self.graphics_card_var.get(),
                self.system_info['os_info'],
                self.special_software_var.get(),
                self.case_appearance_var.get(),
                self.monitor_appearance_var.get(),
                self.keyboard_appearance_var.get(),
                self.printer_appearance_var.get(),
                self.system_info['cpu_model'],
                self.system_info['cpu_count'],
                self.system_info['ram_total_gb'],
                self.system_info['gpu_model'],
                self.system_info['disk_total_gb'],
                self.purchase_date_var.get(),
                self.notes_var.get(),
                self.system_info['collection_date']
            ))

            conn.commit()
            conn.close()

            messagebox.showinfo("موفقیت", "اطلاعات با موفقیت ثبت شد!")
            self.clear_form()

        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره اطلاعات:\n{str(e)}")

    def clear_form(self):
        self.name_var.set("")
        self.location_var.set("")
        self.case_model_var.set("")
        self.monitor_model_var.set("")
        self.keyboard_model_var.set("")
        self.mouse_model_var.set("")
        self.printer_model_var.set("")
        self.internal_phone_var.set("")
        self.delivery_date_var.set("")
        self.recipient_name_var.set("")
        self.switch_connected_var.set("")
        self.motherboard_model_var.set("")
        self.has_dvd_var.set(False)
        self.disk_type_capacity_var.set("")
        self.ram_capacity_gb_var.set("")
        self.graphics_card_var.set("")
        self.special_software_var.set("")
        self.case_appearance_var.set("")
        self.monitor_appearance_var.set("")
        self.keyboard_appearance_var.set("")
        self.printer_appearance_var.set("")
        self.purchase_date_var.set("")
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Text):
                widget.delete("1.0", "end")

    def open_password_window(self):
        password_window = Toplevel(self.root)
        password_window.title("ورود رمز عبور")
        password_window.geometry("300x150")
        password_window.resizable(False, False)

        ttk.Label(password_window, text="رمز عبور:").pack(pady=10)
        password_entry = Entry(password_window, show="*")
        password_entry.pack(pady=5)

        def check_password():
            if password_entry.get() == "admin123":
                self.export_to_excel()
                password_window.destroy()
            else:
                messagebox.showerror("خطا", "رمز عبور اشتباه است!")
                password_window.destroy()

        ttk.Button(password_window, text="تایید", command=check_password).pack(pady=10)

    def export_to_excel(self):
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM devices")
            rows = cursor.fetchall()
            conn.close()

            if not rows:
                messagebox.showinfo("اطلاع", "هیچ داده‌ای برای خروجی وجود ندارد.")
                return

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "لیست دستگاه‌ها"

            headers = [
                "ID", "نام دستگاه", "محل اسکان", "نام کامپیوتر", "مدل کیس", "مدل مانیتور",
                "مدل کیبورد", "مدل ماوس", "مدل چاپگر", "شماره داخلی", "تاریخ تحویل",
                "تحویل گیرنده", "IPv4", "IPv6", "MAC", "سوئیچ متصل", "مدل مادربرد",
                "DVD دارد", "هارد/SSD", "رم (GB)", "کارت گرافیک", "سیستم عامل",
                "نرم‌افزارهای خاص", "ظاهر کیس", "ظاهر مانیتور", "ظاهر کیبورد",
                "ظاهر چاپگر", "پردازنده", "تعداد هسته", "RAM (GB) - خودکار",
                "کل فضای دیسک (GB)", "تاریخ خرید", "یادداشت‌ها", "تاریخ ثبت"
            ]

            ws.append(headers)

            for row in rows:
                ws.append(row)

            # استایل سرصفحه
            for col in ws[1]:
                col.font = Font(bold=True)
                col.alignment = Alignment(horizontal="center")

            filename = f"devices_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            wb.save(filename)

            messagebox.showinfo("موفقیت", f"داده‌ها با موفقیت در فایل {filename} ذخیره شدند.")

        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره فایل اکسل:\n{str(e)}")

# ————————————————————————————————————————————————————————————————
# اجرای برنامه
# ————————————————————————————————————————————————————————————————

if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    app = SystemInventoryBasicApp(root)
    root.mainloop()
