# نمونهٔ خروجی اسکیل

`nemune-list-150-tajrobi-1404.md` و `nemune-list-150-tajrobi-1404.xlsx` خروجی اسکیل [`konkur-entekhab-reshteh`](../.claude/skills/konkur-entekhab-reshteh/SKILL.md) هستند، برای یک داوطلب **فرضی** و روی **اکسل تجربی ۱۴۰۴**. هدفشان فقط نشان دادن شکل خروجی است. کدها مال ۱۴۰۴ هستند و برای ۱۴۰۵ معتبر نیستند.

- **پروفایل ورودی (بازهٔ مشاور):** [`profile-example.json`](../.claude/skills/konkur-entekhab-reshteh/assets/profile-example.json)
- **دستور ساخت:**

```bash
python .claude/skills/konkur-entekhab-reshteh/scripts/build_list.py \
  --booklet انتخاب-رشته-تجربی-1404.xlsx \
  --profile .claude/skills/konkur-entekhab-reshteh/assets/profile-example.json \
  --out-dir examples --name nemune-list-150-tajrobi-1404
```

شیت «خلاصه» فایل اکسل فرمول دارد و وقتی فایل در Excel باز شود محاسبه می‌شود.
