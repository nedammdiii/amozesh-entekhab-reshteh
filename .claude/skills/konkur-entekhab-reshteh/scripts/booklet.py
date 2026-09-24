"""خواندن اکسل استخراج‌شده از دفترچهٔ انتخاب رشتهٔ سنجش و ساختن ستون‌های کمکی.

برای اکسل هر سال (مثلاً ۱۴۰۵) کار می‌کند، حتی اگر نام ستون‌ها، نام شیت یا ردیف عنوان کمی
فرق کند: ستون‌ها با نام‌های مشابه شناخته می‌شوند و هر چیز ناشناخته در df.attrs["warnings"]
گزارش می‌شود. اگر ستونی شناخته نشد، با یک فایل JSON (--column-map) نگاشتش کنید.
قالب مرجع (نمونه‌های ۱۴۰۴): references/rahnama-kamel.md بخش ۱۷.
"""
import json
import re
import sys

import pandas as pd

SHEET = "رشته محل ها"
# نام استاندارد ← نام‌های مشابهی که در فایل‌ها دیده می‌شود (مقایسه بدون فاصله و علائم)
ALIASES = {
    "کد رشته محل": ["کد رشته محل", "کدرشته محل", "کد رشته", "کد", "کد رشته‌محل"],
    "عنوان رشته": ["عنوان رشته", "رشته", "نام رشته", "رشته تحصیلی"],
    "نحوه پذیرش": ["نحوه پذیرش", "نوع پذیرش", "پذیرش"],
    "دوره تحصیلی": ["دوره تحصیلی", "دوره", "نوع دوره"],
    "جنس پذیرش": ["جنس پذیرش", "جنسیت", "جنس"],
    "استان": ["استان", "استان محل تحصیل"],
    "دانشگاه / مؤسسه": ["دانشگاه / مؤسسه", "دانشگاه/موسسه", "دانشگاه", "موسسه", "نام دانشگاه",
                        "نام موسسه", "دانشگاه یا موسسه", "محل تحصیل"],
    "نوع دانشگاه": ["نوع دانشگاه", "نوع موسسه"],
    "زیرمجموعه": ["زیرمجموعه", "وزارتخانه"],
    "ظرفیت نیمسال اول": ["ظرفیت نیمسال اول", "نیمسال اول", "ظرفیت مهر"],
    "ظرفیت نیمسال دوم": ["ظرفیت نیمسال دوم", "نیمسال دوم", "ظرفیت بهمن"],
    "ظرفیت کل": ["ظرفیت کل", "ظرفیت", "جمع ظرفیت"],
    "ظرفیت زن": ["ظرفیت زن"],
    "ظرفیت مرد": ["ظرفیت مرد"],
    "شرط بومی / سهمیه": ["شرط بومی / سهمیه", "شرط بومی", "سهمیه", "بومی"],
    "توضیحات": ["توضیحات", "توضیح", "ملاحظات"],
    "بخش دفترچه": ["بخش دفترچه", "بخش"],
    "منبع دوره تحصیلی": ["منبع دوره تحصیلی"],
    "صفحه PDF": ["صفحه PDF", "صفحه", "شماره صفحه"],
    "شروع": ["شروع", "زمان شروع", "سال شروع", "شروع تحصیل", "زمان شروع تحصیل"],
}
REQUIRED = ["کد رشته محل", "عنوان رشته", "دوره تحصیلی", "دانشگاه / مؤسسه"]
NUMERIC = ["ظرفیت نیمسال اول", "ظرفیت نیمسال دوم", "ظرفیت کل"]
WITH_EXAM = "با آزمون"
RECORDS_ONLY = "صرفا با سوابق تحصیلی"
KNOWN_COURSES = {"روزانه", "روزانه - غیردولتی", "نوبت دوم", "پردیس خودگردان", "شهریه پرداز",
                 "آزاد تمام وقت", "خودگردان آزاد", "غیرانتفاعی", "پیام نور", "مجازی", "مشترک"}
# قبولی در این چهار رشته در هر دوره‌ای = محرومیت از کنکور سال بعد
BAN_FIELDS = {"دکتری عمومی پزشکی", "دکتری عمومی دندانپزشکی", "دکتری عمومی داروسازی",
              "دکتری عمومی دامپزشکی"}
_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def norm(s):
    """یکسان‌سازی متن فارسی: ی/ک عربی، نیم‌فاصله، ارقام، فاصله‌های اضافه، املای رایج PDF."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).translate(_DIGITS).replace("ي", "ی").replace("ك", "ک").replace("‌", " ")
    s = s.replace("کهکیلویه", "کهگیلویه")
    return re.sub(r"\s+", " ", s).strip()


def _key(s):
    return re.sub(r"[\s/\\\-()_.:]", "", norm(s).replace("ؤ", "و").replace("أ", "ا").replace("ـ", "")).lower()


ALIAS_TO_STD = {_key(a): std for std, names in ALIASES.items() for a in [std] + names}


def province_pattern(province):
    """الگوی «استان X» یا «استان محروم X» که «کرمان» را با «کرمانشاه» اشتباه نگیرد."""
    prov = r"\s*".join(map(re.escape, norm(province).split(" ")))
    return rf"استان(?:\s*محروم)?\s*{prov}(?![آ-ی])"


def _find_table(xl):
    """شیت و ردیفی را پیدا می‌کند که بیشترین ستون شناخته‌شده را دارد."""
    best = None
    for sh in xl.sheet_names:
        raw = pd.read_excel(xl, sheet_name=sh, header=None, dtype=str, nrows=15)
        for i, row in raw.iterrows():
            hits = len({ALIAS_TO_STD[_key(v)] for v in row.tolist() if _key(v) in ALIAS_TO_STD})
            score = (hits, sh == SHEET)
            if hits >= 3 and (best is None or score > best[2]):
                best = (sh, i, score)
    return best


def _gender(v):
    v = norm(v)
    if ("زن" in v and "مرد" in v) or any(x in v for x in ("هر دو", "هردو", "مختلط")) or not v:
        return "زن و مرد"
    if "زن" in v:
        return "فقط زن"
    if "مرد" in v:
        return "فقط مرد"
    return v


def load_booklet(path, column_map=None):
    """جدول کدرشته‌محل‌ها را می‌خواند، ستون‌ها را استاندارد و پرچم‌ها را اضافه می‌کند."""
    warns = []
    xl = pd.ExcelFile(path)
    found = _find_table(xl)
    if not found:
        sys.exit(f"در «{path}» جدولی با ستون‌های شناخته‌شده (کد رشته محل، عنوان رشته، دوره، دانشگاه و…) "
                 "پیدا نشد. با --column-map نام ستون‌ها را نگاشت کنید.")
    sheet, header_row, _ = found
    df = pd.read_excel(xl, sheet_name=sheet, header=header_row, dtype=str)
    df = df.dropna(how="all")
    original = [str(c) for c in df.columns]
    rename, used = {}, set()
    user_map = {_key(k): v for k, v in (column_map or {}).items()}
    for c in original:
        std = user_map.get(_key(c)) or ALIAS_TO_STD.get(_key(c))
        if std and std not in used:
            rename[c] = std
            used.add(std)
    df = df.rename(columns=rename)
    unknown_cols = [c for c in original if c not in rename and not c.startswith("Unnamed")]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        sys.exit(f"ستون‌های لازم در «{path}» (شیت «{sheet}») شناخته نشد: {missing}\n"
                 f"ستون‌های فایل: {original}\n"
                 "یک فایل JSON مثل {\"نام ستون در فایل\": \"نام استاندارد\"} بسازید و با --column-map بدهید.")

    for c in df.columns:
        if c not in NUMERIC:
            df[c] = df[c].map(norm)
    if "نحوه پذیرش" not in df.columns:
        df["نحوه پذیرش"] = WITH_EXAM
        warns.append("ستون «نحوه پذیرش» نبود؛ همهٔ ردیف‌ها «با آزمون» فرض شدند.")
    else:
        df["نحوه پذیرش"] = df["نحوه پذیرش"].map(lambda v: RECORDS_ONLY if "سوابق" in v else WITH_EXAM)
    if "جنس پذیرش" not in df.columns:
        df["جنس پذیرش"] = ""
        warns.append("ستون «جنس پذیرش» نبود؛ فیلتر جنسیت عمل نمی‌کند.")
    df["جنس پذیرش"] = df["جنس پذیرش"].map(_gender)
    for c in ["استان", "شرط بومی / سهمیه", "توضیحات"]:
        if c not in df.columns:
            df[c] = ""
            warns.append(f"ستون «{c}» نبود؛ فیلترها و پرچم‌های وابسته به آن کامل نیستند.")
    for c in ["نوع دانشگاه", "زیرمجموعه", "ظرفیت زن", "ظرفیت مرد", "بخش دفترچه", "منبع دوره تحصیلی", "صفحه PDF"]:
        if c not in df.columns:
            df[c] = ""
    for c in NUMERIC:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c].map(norm), errors="coerce").fillna(0).astype(int)
    if "ظرفیت کل" not in df.columns:
        df["ظرفیت کل"] = df.get("ظرفیت نیمسال اول", 0) + df.get("ظرفیت نیمسال دوم", 0)
    if "ظرفیت نیمسال اول" not in df.columns:
        df["ظرفیت نیمسال اول"] = df["ظرفیت کل"]
        df["ظرفیت نیمسال دوم"] = 0
    df["کد رشته محل"] = df["کد رشته محل"].str.replace(r"\.0$", "", regex=True)
    df = df[df["کد رشته محل"].str.fullmatch(r"\d{4,6}")]

    odd = sorted(set(df["دوره تحصیلی"]) - KNOWN_COURSES)
    if odd:
        warns.append(f"مقدارهای جدید در «دوره تحصیلی»: {odd}. در allowed_courses و choices دقیقاً همین‌ها را "
                     "بنویسید؛ دوره‌هایی که با «روزانه» شروع می‌شوند رایگان فرض شده‌اند.")
    df["_title"] = df["عنوان رشته"]
    df["_uni"] = df["دانشگاه / مؤسسه"]
    df["_notes"] = df["توضیحات"]
    df["_quota"] = df["شرط بومی / سهمیه"]
    df["_section"] = df["بخش دفترچه"]
    df["_source"] = str(path)
    df = add_flags(df)
    df.attrs["warnings"] = warns
    df.attrs["mapping"] = {"sheet": sheet, "header_row": int(header_row) + 1, "columns": rename,
                           "unrecognized": unknown_cols}
    return df


def add_flags(df):
    """پرچم‌های بخش ۱۷ راهنما. خوابگاه عمداً پرچم نمی‌شود (عبارت‌های خوابگاه فقط عدم تضمین‌اند)."""
    n, q, sec = df["_notes"], df["_quota"], df["_section"]
    df["تعهد خدمت"] = n.str.contains("تعهد خدمت") | sec.str.contains("تعهد خدمت|افزایش ظرفیت")
    df["مصاحبه/شرایط خاص"] = n.str.contains("مصاحبه|پیوستها|گزینش|بورس|معاینات")
    df["ویژه بهیاران"] = n.str.contains("ویژه بهیاران")
    df["شهریه‌ای"] = ~df["دوره تحصیلی"].str.startswith("روزانه")
    start = pd.Series("مهر", index=df.index)
    start[(df["ظرفیت نیمسال اول"] == 0) & (df["ظرفیت نیمسال دوم"] > 0)] = "بهمن"
    text = n + " " + df["دوره تحصیلی"] + " " + df["_title"] + " " + df["_uni"] + " " + sec
    if "شروع" in df.columns:
        s = df["شروع"].fillna("")
        start[s.str.contains("بهمن")] = "بهمن"
        text = text + " " + s
    start[text.str.contains("1406")] = "مهر ۱۴۰۶"
    df["شروع"] = start
    df["ظرفیت کم"] = df["ظرفیت کل"] <= 2
    df["محرومیت کنکور بعد"] = ~df["شهریه‌ای"] | df["_title"].isin(BAN_FIELDS)
    df["فقط بومی"] = q.str.contains("مخصوص")
    df["سهمیه خاص"] = df["فقط بومی"] & q.str.contains("بلایای طبیعی|شهرستان")
    return df


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)
