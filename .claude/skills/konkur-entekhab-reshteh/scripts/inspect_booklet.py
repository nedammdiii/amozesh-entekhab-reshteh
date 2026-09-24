"""وارسی سریع اکسل دفترچه: ساختار، آمار، نام دقیق رشته‌ها و جستجو.

نمونه:
    python inspect_booklet.py booklet.xlsx                 # خلاصهٔ فایل
    python inspect_booklet.py booklet.xlsx --titles        # همهٔ عنوان‌های «با آزمون» با تعداد کد
    python inspect_booklet.py booklet.xlsx --search فیزیو  # جستجو در عنوان رشته و نام دانشگاه
    python inspect_booklet.py booklet.xlsx --code 33194    # جزئیات یک کد
"""
import argparse

import pandas as pd

from booklet import WITH_EXAM, load_booklet, load_json, norm

p = argparse.ArgumentParser(description="وارسی اکسل دفترچهٔ انتخاب رشته")
p.add_argument("booklet")
p.add_argument("--titles", action="store_true", help="فهرست عنوان‌های رشته (فقط با آزمون)")
p.add_argument("--search", help="متن برای جستجو در عنوان رشته و دانشگاه (فقط با آزمون)")
p.add_argument("--code", help="نمایش کامل یک کد رشته محل")
p.add_argument("--column-map", help='JSON نگاشت نام ستون‌ها، مثل {"نام ستون در فایل": "کد رشته محل"}')
a = p.parse_args()

df = load_booklet(a.booklet, load_json(a.column_map) if a.column_map else None)
exam = df[df["نحوه پذیرش"] == WITH_EXAM]
pd.set_option("display.width", 200)
pd.set_option("display.max_colwidth", 70)
pd.set_option("display.max_rows", 500)

if a.code:
    row = df[df["کد رشته محل"] == a.code.strip()]
    if row.empty:
        print("کد پیدا نشد.")
    for _, r in row.iterrows():
        for c in [c for c in df.columns if not c.startswith("_")]:
            print(f"{c}: {r[c]}")
elif a.search:
    s = norm(a.search)
    hit = exam[exam["_title"].str.contains(s, regex=False) | exam["_uni"].str.contains(s, regex=False)]
    print(f"{len(hit)} کد «با آزمون» پیدا شد.\n")
    print(hit["عنوان رشته"].value_counts().to_string())
    print()
    print(hit[["کد رشته محل", "عنوان رشته", "دانشگاه / مؤسسه", "دوره تحصیلی", "جنس پذیرش", "ظرفیت کل"]]
          .head(60).to_string(index=False))
elif a.titles:
    print(exam["عنوان رشته"].value_counts().to_string())
else:
    mp = df.attrs.get("mapping", {})
    print(f"فایل: {a.booklet}")
    print(f"شیت: «{mp.get('sheet')}» | ردیف عنوان ستون‌ها: {mp.get('header_row')}")
    renamed = {k: v for k, v in mp.get("columns", {}).items() if k != v}
    if renamed:
        print("ستون‌هایی که با نام مشابه شناخته شدند:", renamed)
    if mp.get("unrecognized"):
        print("ستون‌های ناشناخته (استفاده نمی‌شوند):", mp["unrecognized"])
    for w in df.attrs.get("warnings", []):
        print("هشدار:", w)
    print(f"کل ردیف‌ها: {len(df):,} | «با آزمون» (فرم ۱۵۰تایی): {len(exam):,} | "
          f"«صرفا با سوابق تحصیلی» (ثبت‌نام جدا): {len(df) - len(exam):,}")
    print(f"بازهٔ کدهای با آزمون: {exam['کد رشته محل'].min()} تا {exam['کد رشته محل'].max()}")
    print(f"تعداد عنوان رشته (با آزمون): {exam['عنوان رشته'].nunique()}")
    for col in ["دوره تحصیلی", "جنس پذیرش", "نوع دانشگاه"]:
        print(f"\n{col} (با آزمون):")
        print(exam[col].value_counts().to_string())
    flags = ["تعهد خدمت", "مصاحبه/شرایط خاص", "ویژه بهیاران", "فقط بومی", "سهمیه خاص", "ظرفیت کم"]
    print("\nپرچم‌ها (با آزمون):")
    print(exam[flags].sum().to_string())
    print("\nشروع (با آزمون):")
    print(exam["شروع"].value_counts().to_string())
