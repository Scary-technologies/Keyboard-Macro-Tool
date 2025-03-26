import openpyxl
import keyboard
import pyautogui
import time

# تابع برای بارگذاری دستورات از فایل اکسل
def load_commands(filename):
    try:
        wb_obj = openpyxl.load_workbook(filename)
        sheet_obj = wb_obj.active
        commands = {}
        for row in sheet_obj.iter_rows(min_row=2, values_only=True):
            key = row[0]  # ستون اول: کلید
            action = row[1]  # ستون دوم: اقدام
            action_type = row[2] if len(row) > 2 else "write"  # ستون سوم: نوع اقدام (پیش‌فرض write)
            if key:  # اگه کلید وجود داشته باشه
                commands[key] = (action, action_type)
        return commands
    except Exception as e:
        print(f"خطا در بارگذاری فایل اکسل: {e}")
        return {}

# تنظیمات اولیه
filename = "Commands.xlsx"  # می‌تونی این رو به یه مسیر دلخواه تغییر بدی
commands = load_commands(filename)

# حلقه اصلی
while True:
    try:
        x = keyboard.read_key()  # خوندن کلید فشرده شده
        if x in commands:  # اگه کلید توی دیکشنری باشه
            action, action_type = commands[x]
            while keyboard.is_pressed(x):  # تا وقتی کلید نگه داشته شده
                if action_type == "write" and action:
                    keyboard.write(action)  # تایپ کردن متن
                elif action_type == "click":
                    pyautogui.click(pyautogui.position())  # کلیک در موقعیت فعلی ماوس
                elif action_type == "hotkey" and action:
                    keyboard.send(action)  # اجرای کلید ترکیبی
                time.sleep(0.1)  # تاخیر برای جلوگیری از اجرای بیش از حد
    except Exception as e:
        print(f"خطا: {e}")