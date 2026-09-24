# amozesh-entekhab-reshteh

اسکیل (Agent Skill) برای ساختن لیست ۱۵۰تایی انتخاب رشتهٔ کنکور سراسری. ورودی‌هایش اکسل دفترچهٔ همان سال و جواب‌های دانش‌آموز به فرم تلگرام است.

- **اسکیل:** [`.claude/skills/konkur-entekhab-reshteh/`](.claude/skills/konkur-entekhab-reshteh/)
  - [`SKILL.md`](.claude/skills/konkur-entekhab-reshteh/SKILL.md): روند کار برای ایجنت
  - [`references/rahnama-kamel.md`](.claude/skills/konkur-entekhab-reshteh/references/rahnama-kamel.md): راهنمای کامل انتخاب رشته (سهمیه‌ها، بومی‌گزینی، دوره‌ها، محرومیت‌ها، ساختار اکسل دفترچه)
  - [`scripts/`](.claude/skills/konkur-entekhab-reshteh/scripts/): خواندن و وارسی اکسل (`inspect_booklet.py`) و ساخت لیست (`build_list.py`)
  - [`assets/telegram-form.md`](.claude/skills/konkur-entekhab-reshteh/assets/telegram-form.md): فرم تلگرام دانش‌آموز (برای کپی) و جدول تبدیل جواب‌ها
  - [`assets/`](.claude/skills/konkur-entekhab-reshteh/assets/): کلیدهای پروفایل و نمونهٔ پروفایل
- **فایل آماده برای نصب:** [`dist/konkur-entekhab-reshteh.skill`](dist/konkur-entekhab-reshteh.skill)
- **نمونهٔ خروجی (دادهٔ ۱۴۰۴، داوطلب فرضی):** [`examples/`](examples/)

## استفاده

1. نصب:
   - **Claude (claude.ai، اپ دسکتاپ و موبایل):** فایل `dist/konkur-entekhab-reshteh.skill` را در Settings ← Capabilities ← Skills بارگذاری کنید، یا روی دکمهٔ «Save skill» کارت فایل در گفتگو بزنید.
   - **Claude Code روی همین مخزن:** نیازی به کاری نیست؛ اسکیل در `.claude/skills/` است و خودکار بارگذاری می‌شود.
   - **ایجنت‌های دیگر:** پوشهٔ `.claude/skills/konkur-entekhab-reshteh` را در پوشهٔ اسکیل‌های ایجنت بگذارید.
2. فرم تلگرام را برای دانش‌آموز بفرستید. بعد اکسل دفترچهٔ ۱۴۰۵ گروه او را همراه با جواب‌هایش به ایجنت بدهید (پرامپت آماده در پیوست ج راهنما).
3. ایجنت اکسل را وارسی می‌کند، پروفایل را می‌سازد، لیست را به‌صورت xlsx و md تحویل می‌دهد و هشدارها را فهرست می‌کند.

اسکریپت‌ها به Python 3 و کتابخانه‌های `pandas` و `openpyxl` نیاز دارند.
