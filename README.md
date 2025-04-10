برای دانلود برنامه به صورت فایل زیپ میتوانید از این لینک استفاده کنید 
**[Ver1.0.0](https://github.com/Scary-technologies/Keyboard-Macro-Tool/archive/refs/tags/Macro.zip)
## پروژه ابزار ماکرو کیبورد (Keyboard Macro Tool)

این پروژه یک ابزار گرافیکی ساده برای اجرای خودکار دستورات با استفاده از کلیدهای کیبورد است. دستورات و عملکردها از یک فایل اکسل خوانده می‌شوند و با فشردن کلیدهای خاص، اقدامات مشخص‌شده به‌صورت خودکار انجام می‌گیرد.

از کتابخانه‌های زیر استفاده شده است:
- `keyboard` برای دریافت ورودی‌های کیبورد و ثبت میانبرها  
- `openpyxl` برای خواندن فایل اکسل  
- `pyautogui` برای اجرای عملکردهایی مانند تایپ، کلیک و کلیدهای ترکیبی  
- `tkinter` برای رابط گرافیکی ساده

---

### نصب و راه‌اندازی

برای اجرای این پروژه، نیاز به پایتون 3 و نصب کتابخانه‌های زیر دارید:

```bash
pip install openpyxl keyboard pyautogui
```

---

### نحوه استفاده

1. **ساخت فایل اکسل**  
   - نام فایل: `Commands.xlsx`
   - باید در همان پوشه‌ای باشد که فایل `KeyboardMacroTool.py` قرار دارد.
   - اطلاعات باید به صورت زیر وارد شوند:

| کلید (A) | عملکرد (B) | نوع اقدام (C) |
|----------|------------|----------------|
| a        | hello      | write          |
| b        |            | click          |
| c        | ctrl+v     | hotkey         |

2. **اجرای برنامه**  
   فایل `KeyboardMacroTool.py` را اجرا کنید:

```bash
python KeyboardMacroTool.py
```

با اجرای برنامه، رابط گرافیکی کوچکی باز می‌شود که نشان می‌دهد برنامه فعال است و می‌توانید با فشردن کلیدهای تعیین‌شده، عملیات مورد نظر را انجام دهید.

---

### عملکرد ستون‌های فایل اکسل

- **ستون A (کلید):** کلیدی که با فشردن آن، عملکرد اجرا می‌شود.
- **ستون B (عملکرد):**  
  - برای `write`: متنی که باید تایپ شود.  
  - برای `click`: خالی بگذارید (کلیک در موقعیت فعلی ماوس).  
  - برای `hotkey`: ترکیب کلیدها مثل `ctrl+v` یا `alt+tab`.

- **ستون C (نوع اقدام):** یکی از این مقادیر:
  - `write`: تایپ متن
  - `click`: کلیک کردن ماوس
  - `hotkey`: اجرای کلیدهای ترکیبی

---

### نکات مهم

- **اجرا با نگه داشتن کلید:** اگر کلید را نگه دارید، عملکرد با تأخیر 0.1 ثانیه تکرار می‌شود (تا از اجرای بیش‌ازحد جلوگیری شود).
- **رفع باگ KeyError:** نسخه جدید از `add_hotkey` و `remove_hotkey` با `suppress=True` استفاده می‌کند تا از خطاهای مربوط به `unhook/unblock_key` جلوگیری شود.
- **خروج از برنامه:** می‌توانید با بستن پنجره یا زدن `Ctrl+C` در ترمینال، برنامه را ببندید.
- **لاگ‌ها:** پیام‌ها و خطاهای اجرایی در کنسول چاپ می‌شوند.

---

### توسعه و سفارشی‌سازی

می‌توانید کد را با توجه به نیاز خود تغییر دهید:

- اضافه کردن عملکردهای جدید (مثلاً باز کردن برنامه، اسکرین‌شات، حرکت ماوس و...)
- تعریف کلید خاص برای توقف ماکرو (مثلاً ESC)
- تغییر نام یا مسیر فایل اکسل از طریق متغیر `filename`

مثال: افزودن کلید "esc" برای خروج از برنامه

```python
keyboard.add_hotkey("esc", lambda: exit(0), suppress=True)
```

---

این ابزار می‌تواند در اتوماسیون کارهای روزمره، پاسخ سریع به دستورات، یا کمک به افراد با نیازهای خاص بسیار مفید باشد.

---  
هرگونه پیشنهاد یا توسعه‌ای دارید، خوشحال می‌شوم در گیت‌هاب بررسی کنم!
