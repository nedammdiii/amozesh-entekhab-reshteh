# amozesh-entekhab-reshteh

اسکیل (Agent Skill) برای ساختن لیست ۱۵۰تایی انتخاب رشتهٔ کنکور سراسری. ورودی‌هایش اکسل دفترچهٔ همان سال، اطلاعات داوطلب و بازهٔ «خوشبینانه / واقع‌بینانه / بدبینانه» است که مشاور تعیین می‌کند.

- **اسکیل:** [`skills/konkur-entekhab-reshteh/`](skills/konkur-entekhab-reshteh/)
  - [`SKILL.md`](skills/konkur-entekhab-reshteh/SKILL.md): روند کار برای ایجنت
  - [`references/rahnama-kamel.md`](skills/konkur-entekhab-reshteh/references/rahnama-kamel.md): راهنمای کامل انتخاب رشته (سهمیه‌ها، بومی‌گزینی، دوره‌ها، محرومیت‌ها، ساختار اکسل دفترچه)
  - [`scripts/`](skills/konkur-entekhab-reshteh/scripts/): خواندن و وارسی اکسل (`inspect_booklet.py`) و ساخت لیست (`build_list.py`)
  - [`assets/`](skills/konkur-entekhab-reshteh/assets/): فرم ورودی مشاور و نمونهٔ پروفایل
- **فایل آماده برای نصب:** [`dist/konkur-entekhab-reshteh.skill`](dist/konkur-entekhab-reshteh.skill)
- **نمونهٔ خروجی (دادهٔ ۱۴۰۴، داوطلب فرضی):** [`examples/`](examples/)

## استفاده

1. فایل `.skill` را در Claude نصب کنید، یا پوشهٔ `skills/konkur-entekhab-reshteh` را در پوشهٔ اسکیل‌های ایجنت خود بگذارید.
2. اکسل دفترچهٔ ۱۴۰۵ گروه داوطلب را به ایجنت بدهید، همراه با مشخصات داوطلب و بازهٔ مشاور (پرامپت آماده در پیوست ج راهنما).
3. ایجنت اکسل را وارسی می‌کند، پروفایل را می‌سازد، لیست را به‌صورت xlsx و md تحویل می‌دهد و هشدارها را فهرست می‌کند.

اسکریپت‌ها به Python 3 و کتابخانه‌های `pandas` و `openpyxl` نیاز دارند.
