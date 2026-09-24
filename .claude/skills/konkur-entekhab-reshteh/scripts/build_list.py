"""ساخت لیست حداکثر ۱۵۰تایی انتخاب رشته از اکسل دفترچه + پروفایل داوطلب (JSON).

بازهٔ «خوشبینانه / واقع‌بینانه / بدبینانه» را مشاور تعیین می‌کند و در پروفایل می‌آید؛
این اسکریپت شانس قبولی را تخمین نمی‌زند. ترتیب لیست فقط بر اساس ترجیح داوطلب است.

نمونه:
    python build_list.py --booklet tajrobi-1405.xlsx --profile profile.json --out-dir out --name sara
    python build_list.py --booklet tajrobi.xlsx --booklet zaban.xlsx --profile p.json   # چند گروه

قالب پروفایل: assets/profile-example.json و assets/intake-form.md
"""
import argparse
import json
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from booklet import WITH_EXAM, load_booklet, load_json, norm, province_pattern

TIERS = ["خوشبینانه", "واقع‌بینانه", "بدبینانه"]
DEFAULT_CAPS = {"خوشبینانه": 40, "واقع‌بینانه": 80, "بدبینانه": 30}
REFILL_ORDER = ["واقع‌بینانه", "بدبینانه", "خوشبینانه"]  # جای خالی یک سطح اول به این‌ها می‌رسد

# سطح‌بندی پیش‌فرض دانشگاه‌ها (فقط برای ترتیب درون یک ورودی؛ پروفایل می‌تواند اضافه کند)
UNI_TIER1 = ["علوم پزشکی تهران", "شهید بهشتی", "علوم پزشکی ایران", "علوم پزشکی شیراز",
             "علوم پزشکی اصفهان", "علوم پزشکی مشهد", "علوم پزشکی تبریز", "علوم توانبخشی",
             "صنعتی شریف", "دانشگاه تهران", "صنعتی امیرکبیر", "علم و صنعت", "صنعتی اصفهان",
             "خواجه نصیر", "علامه طباطبایی"]
UNI_TIER3 = ["دانشکده", "پیام نور", "غیرانتفاعی", "ملی مهارت", "گناباد", "دزفول", "تربت حیدریه",
             "شاهرود", "سبزوار", "زابل", "جهرم", "فسا", "آبادان", "خراسان شمالی", "نیشابور",
             "ایرانشهر", "جیرفت", "علوم پزشکی بم", "یاسوج"]

FLAG_COLS = ["تعهد خدمت", "شهریه‌ای", "مصاحبه/شرایط خاص", "فقط بومی", "ظرفیت کم", "محرومیت کنکور بعد"]
OUT_COLS = ["اولویت", "کد رشته محل", "عنوان رشته", "دانشگاه / مؤسسه", "استان", "دوره تحصیلی",
            "جنس پذیرش", "شروع", "ظرفیت کل", "سطح (مشاور)"] + FLAG_COLS + [
            "شرط بومی / سهمیه", "توضیحات", "صفحه PDF"]


def tier_name(x):
    t = norm(x).replace(" ", "")
    for prefix, name in (("خوش", TIERS[0]), ("واقع", TIERS[1]), ("بد", TIERS[2])):
        if t.startswith(prefix):
            return name
    raise SystemExit(f"سطح نامعتبر در پروفایل: «{x}» (یکی از {TIERS})")


def as_list(v):
    if v is None or v is False:
        return []
    return v if isinstance(v, list) else [v]


def entry_mask(df, e, allowed_courses):
    m = pd.Series(True, index=df.index)
    if e.get("title"):
        m &= df["_title"] == norm(e["title"])
    if e.get("title_contains"):
        m &= df["_title"].str.contains(norm(e["title_contains"]), regex=False)
    courses = as_list(e.get("courses")) or allowed_courses
    m &= df["دوره تحصیلی"].isin(courses)
    unis = [norm(u) for u in as_list(e.get("universities"))]
    if unis:
        m &= df["_uni"].map(lambda u: any(x in u for x in unis))
    ex = [norm(u) for u in as_list(e.get("exclude_universities"))]
    if ex:
        m &= ~df["_uni"].map(lambda u: any(x in u for x in ex))
    provs = as_list(e.get("provinces"))
    if provs:
        m &= df["استان"].isin(provs)
    starts = as_list(e.get("start"))
    if starts:
        m &= df["شروع"].isin(starts)
    return m


def uni_tier(uni, extra):
    if "(محل تحصیل" in uni or "(دانشکده" in uni:
        return 3
    if any(x in uni for x in UNI_TIER1 + extra.get("1", [])):
        return 1
    if any(x in uni for x in UNI_TIER3 + extra.get("3", [])):
        return 3
    return 2


def rank_score(names, value):
    """امتیاز ۰ تا ۳ بر اساس جایگاه در یک فهرست ترتیبی (اولی = ۳)."""
    for i, n in enumerate(names):
        if n and n in value:
            return 3 * (1 - i / max(len(names), 1))
    return None


def main():
    ap = argparse.ArgumentParser(description="ساخت لیست ۱۵۰تایی انتخاب رشته")
    ap.add_argument("--booklet", action="append", required=True, help="اکسل دفترچه (برای چند گروه تکرار شود)")
    ap.add_argument("--profile", required=True, help="پروفایل داوطلب (JSON)")
    ap.add_argument("--out-dir", default=".")
    ap.add_argument("--name", default="list", help="پیشوند نام فایل‌های خروجی")
    ap.add_argument("--column-map", help='JSON نگاشت نام ستون‌ها، مثل {"نام ستون در فایل": "کد رشته محل"}')
    a = ap.parse_args()

    prof = load_json(a.profile)
    cmap = load_json(a.column_map) if a.column_map else None
    frames = [load_booklet(b, cmap) for b in a.booklet]
    booklet_warns = [f"{Path(b).name}: {w}" for b, f in zip(a.booklet, frames)
                     for w in f.attrs.get("warnings", [])]
    df = pd.concat(frames, ignore_index=True)
    steps = [("کل ردیف‌های دفترچه", len(df))]

    def step(mask, text):
        nonlocal df
        df = df[mask]
        steps.append((text, len(df)))

    step(df["نحوه پذیرش"] == WITH_EXAM, "فقط «با آزمون» (کدهای فرم ۱۵۰تایی)")
    step(~df["کد رشته محل"].duplicated(), "حذف کدهای تکراری بین گروه‌ها")

    allowed = as_list(prof.get("allowed_courses")) or sorted(df["دوره تحصیلی"].unique())
    choices = prof.get("choices") or []
    if not choices:
        raise SystemExit("پروفایل «choices» ندارد؛ رشته‌ها و بازهٔ مشاور را وارد کنید.")

    # تطبیق هر کد با اولین ورودی سازگار از بازهٔ مشاور
    df = df.copy()
    df["_choice"] = -1
    entry_hits = []
    for i, e in enumerate(choices):
        m = entry_mask(df, e, allowed) & (df["_choice"] == -1)
        df.loc[m, "_choice"] = i
        entry_hits.append(int(m.sum()))
    step(df["_choice"] >= 0, "در بازهٔ مشاور (رشته، دانشگاه، دوره، استان)")
    df["سطح (مشاور)"] = df["_choice"].map(lambda i: tier_name(choices[i].get("tier", "واقع‌بینانه")))
    df["_will_enroll"] = df["_choice"].map(lambda i: bool(choices[i].get("will_enroll", True)))

    gender = prof.get("gender", "")
    if gender:
        step(df["جنس پذیرش"].isin(["زن و مرد", f"فقط {gender}"]), "جنسیت سازگار")
    if not prof.get("is_behyar"):
        step(~df["ویژه بهیاران"], "حذف «ویژه بهیاران»")
    if not prof.get("accept_interview"):
        step(~df["مصاحبه/شرایط خاص"], "حذف مصاحبه، بورس و گزینش")
    native = prof.get("native_province", "")
    own = df["_quota"].str.contains(province_pattern(native), regex=True) if native else False
    ok_native = ~df["فقط بومی"] | (own & (~df["سهمیه خاص"] | bool(prof.get("special_native_quota"))))
    step(ok_native, "حذف کدهای «مخصوص بومی» که داوطلب شرایطش را ندارد")
    commit = prof.get("accept_commitment", False)
    if commit is not True:
        step(~df["تعهد خدمت"] | df["استان"].isin(as_list(commit)), "تعهد خدمت فقط با رضایت (و در استان‌های مجاز)")
    if prof.get("excluded_provinces"):
        step(~df["استان"].isin(prof["excluded_provinces"]), "حذف استان‌های ناخواسته")
    if prof.get("exclude_codes"):
        step(~df["کد رشته محل"].isin([str(c) for c in prof["exclude_codes"]]), "حذف کدهای دستی")
    if prof.get("retake_next_year"):
        step(~df["محرومیت کنکور بعد"] | df["_will_enroll"],
             "حذف کدهای دارای محرومیت که داوطلب در آن‌ها ثبت‌نام نمی‌کند")

    # ترتیب: فقط ترجیح داوطلب (ورودی‌ها به ترتیب علاقه + مکان و دانشگاه)
    extra_tiers = {k: [norm(x) for x in v] for k, v in (prof.get("university_tiers") or {}).items()}
    prov_pref = [str(x) for x in as_list(prof.get("preferred_provinces"))]
    uni_pref = [norm(x) for x in as_list(prof.get("preferred_universities"))]
    home = norm(prof.get("home_city", ""))

    def place(r):
        s = rank_score(prov_pref, r["استان"]) or 0.0
        u = rank_score(uni_pref, r["_uni"])
        s += u if u is not None else {1: 2.0, 2: 1.0, 3: 0.0}[uni_tier(r["_uni"], extra_tiers)]
        if home and home in r["_uni"] and "(" not in r["_uni"]:
            s += 1
        s -= 1.0 * r["شهریه‌ای"] + 1.0 * r["تعهد خدمت"] + 0.3 * (r["شروع"] != "مهر")
        return round(s, 2)

    df["_place"] = df.apply(place, axis=1) if len(df) else []
    mode = prof.get("priority", "field")
    if mode == "place":
        df = df.sort_values(["_place", "_choice", "کد رشته محل"], ascending=[False, True, True])
    elif mode == "balanced":
        df["_key"] = df["_choice"] - 0.5 * df["_place"]
        df = df.sort_values(["_key", "_choice", "کد رشته محل"])
    else:
        df = df.sort_values(["_choice", "_place", "کد رشته محل"], ascending=[True, False, True])
    df["_rank"] = range(len(df))

    # سقف اختیاری هر ورودی («max»): فقط بهترین کدهای آن ورودی (بر اساس ترجیح) می‌مانند
    capped = pd.Series(False, index=df.index)
    for i, e in enumerate(choices):
        if e.get("max") is not None:
            idx = df.index[df["_choice"] == i]
            capped[idx[int(e["max"]):]] = True
    if capped.any():
        step(~capped, "اعمال سقف تعداد کد هر ورودی (max)")

    # انتخاب حداکثر max_codes با سقف هر سطح (جای خالی به سطوح دیگر می‌رسد)
    max_codes = int(prof.get("max_codes", 150))
    caps = {TIERS[i]: v for i, v in enumerate(DEFAULT_CAPS.values())}
    caps.update({tier_name(k): int(v) for k, v in (prof.get("tier_caps") or {}).items()})
    total = sum(caps.values()) or 1
    caps = {k: round(v * max_codes / total) for k, v in caps.items()}
    if len(df) <= max_codes:
        final = df
    else:
        picked = []
        for t in TIERS:
            picked.append(df[df["سطح (مشاور)"] == t].head(caps.get(t, 0)))
        final = pd.concat(picked)
        for t in REFILL_ORDER:
            if len(final) >= max_codes:
                break
            rest = df.drop(final.index)
            final = pd.concat([final, rest[rest["سطح (مشاور)"] == t].head(max_codes - len(final))])
        final = final.sort_values("_rank")
    final = final.head(max_codes).copy()
    final.insert(0, "اولویت", range(1, len(final) + 1))
    dropped = df.drop(final.index)["سطح (مشاور)"].value_counts().to_dict()

    # هشدارها
    warns = list(booklet_warns)
    for i, (e, n) in enumerate(zip(choices, entry_hits)):
        kept = int((final["_choice"] == i).sum())
        label = e.get("title") or e.get("title_contains")
        if n == 0:
            warns.append(f"ورودی {i + 1} («{label}») با هیچ کدی جور نشد؛ نام دقیق را با "
                         f"inspect_booklet.py --search پیدا کنید یا شرط‌هایش را بازتر کنید.")
        elif kept == 0:
            warns.append(f"ورودی {i + 1} («{label}») {n} کد داشت ولی همه با فیلترها یا سقف حذف شدند.")
    tiers_now = final["سطح (مشاور)"].value_counts()
    min_safe = int(prof.get("min_pessimistic", 15))
    if tiers_now.get("بدبینانه", 0) < min_safe:
        warns.append(f"فقط {tiers_now.get('بدبینانه', 0)} کد «بدبینانه» در لیست است (کمتر از {min_safe}). "
                     "اگر داوطلب می‌خواهد حتماً امسال قبول شود، از مشاور رشته‌های بدبینانهٔ بیشتری بگیرید.")
    if len(final) < max_codes:
        warns.append(f"لیست {len(final)} کد دارد (کمتر از {max_codes}). مشکلی نیست، ولی برای پوشش بیشتر می‌شود "
                     "رشته یا دانشگاه‌های بدبینانهٔ دیگری اضافه کرد.")
    if dropped:
        warns.append("به خاطر سقف لیست حذف شد: " + "، ".join(f"{k}: {v}" for k, v in dropped.items())
                     + " (کمترین ترجیح هر سطح حذف شده است).")
    counts = {c: int(final[c].sum()) for c in FLAG_COLS}
    info = [
        (counts["محرومیت کنکور بعد"], "کد محرومیت از کنکور سال بعد دارند (روزانه، یا پزشکی/دندان/دارو/دامپزشکی در هر دوره)."),
        (counts["تعهد خدمت"], "کد تعهد خدمت دارند."),
        (counts["شهریه‌ای"], "کد شهریه‌ای‌اند."),
        (counts["مصاحبه/شرایط خاص"], "کد مصاحبه/گزینش/بورس دارند."),
        (counts["فقط بومی"], "کد «مخصوص بومی» هستند؛ شرط بومی بودن داوطلب را تأیید کنید."),
        (counts["ظرفیت کم"], "کد ظرفیت ۱ یا ۲ نفر دارند."),
        (int((final["شروع"] != "مهر").sum()), "کد شروع غیر از مهر دارند (بهمن یا مهر ۱۴۰۶)."),
    ]
    warns += [f"{n} {t}" for n, t in info if n]

    out = Path(a.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    write_xlsx(out / f"{a.name}.xlsx", final, prof, steps, warns)
    write_md(out / f"{a.name}.md", final, prof, steps, warns, a.name)

    print("\n".join(f"{t}: {n:,}" for t, n in steps))
    print("سطح‌ها در لیست نهایی:", final["سطح (مشاور)"].value_counts().to_dict())
    print("کدها در لیست نهایی:", len(final))
    print("\nهشدارها:\n- " + "\n- ".join(warns))
    print("\nخروجی:", out / f"{a.name}.xlsx", out / f"{a.name}.md")


def cell_value(v):
    if isinstance(v, (bool,)) or type(v).__name__ == "bool_":
        return "✓" if v else ""
    if hasattr(v, "item"):
        v = v.item()
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else v


def write_xlsx(path, final, prof, steps, warns):
    font, bold = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
    head = PatternFill("solid", fgColor="D9D9D9")
    fills = {"خوشبینانه": "FCE5CD", "واقع‌بینانه": "FFF2CC", "بدبینانه": "D9EAD3"}
    wb = Workbook()

    def sheet(title, first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        ws.sheet_view.rightToLeft = True
        return ws

    def put(ws, r, vals, is_head=False):
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = bold if is_head else font
            cell.alignment = Alignment(horizontal="right", vertical="top",
                                       wrap_text=isinstance(v, str) and len(v) > 45)
            if is_head:
                cell.fill = head

    ws = sheet("لیست", first=True)
    put(ws, 1, OUT_COLS, True)
    tier_col = OUT_COLS.index("سطح (مشاور)") + 1
    for r, (_, row) in enumerate(final.iterrows(), 2):
        put(ws, r, [cell_value(row[c]) for c in OUT_COLS])
        ws.cell(row=r, column=tier_col).fill = PatternFill("solid", fgColor=fills[row["سطح (مشاور)"]])
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(OUT_COLS))}{max(len(final), 1) + 1}"
    for i, w in enumerate([7, 11, 24, 42, 15, 13, 9, 9, 8, 12, 9, 8, 12, 8, 8, 12, 40, 55, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws = sheet("خلاصه")
    last = len(final) + 1
    col = {c: get_column_letter(OUT_COLS.index(c) + 1) for c in OUT_COLS}
    rng = lambda c: f"'لیست'!${col[c]}$2:${col[c]}${last}"  # noqa: E731
    r = 1
    put(ws, r, ["سطح (مشاور)", "تعداد"], True)
    for t in TIERS:
        r += 1
        put(ws, r, [t, f"=COUNTIF({rng('سطح (مشاور)')},A{r})"])
    r += 1
    put(ws, r, ["جمع", f"=SUM(B{r - 3}:B{r - 1})"], True)
    r += 2
    put(ws, r, ["رشته", "تعداد"], True)
    for t in final["عنوان رشته"].drop_duplicates():
        r += 1
        put(ws, r, [t, f"=COUNTIF({rng('عنوان رشته')},A{r})"])
    r += 2
    put(ws, r, ["دوره تحصیلی", "تعداد"], True)
    for t in final["دوره تحصیلی"].drop_duplicates():
        r += 1
        put(ws, r, [t, f"=COUNTIF({rng('دوره تحصیلی')},A{r})"])
    r += 2
    put(ws, r, ["پرچم", "تعداد"], True)
    for f in FLAG_COLS:
        r += 1
        put(ws, r, [f, f'=COUNTIF({rng(f)},"✓")'])
    r += 2
    put(ws, r, ["مراحل فیلتر (خروجی build_list.py)", "ردیف باقی‌مانده"], True)
    for t, n in steps:
        r += 1
        put(ws, r, [t, n])
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 16

    ws = sheet("ورودی داوطلب")
    put(ws, 1, ["کلید", "مقدار"], True)
    for i, (k, v) in enumerate(prof.items(), 2):
        put(ws, i, [k, json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else str(v)])
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 120

    ws = sheet("هشدارها")
    put(ws, 1, ["هشدار / نکته"], True)
    for i, w in enumerate(warns + ["کدها را پیش از ثبت با دفترچهٔ رسمی و اطلاعیه‌های اصلاحی سنجش تطبیق دهید."], 2):
        put(ws, i, [w])
        ws.cell(row=i, column=1).alignment = Alignment(horizontal="right", wrap_text=True)
    ws.column_dimensions["A"].width = 120

    wb.calculation.fullCalcOnLoad = True  # فرمول‌های «خلاصه» هنگام باز شدن در Excel محاسبه می‌شوند
    wb.save(path)


def write_md(path, final, prof, steps, warns, name):
    L = []
    w = L.append
    w(f"# لیست انتخاب رشته — {prof.get('student_name', name)}\n")
    w(f"- تعداد کد: **{len(final)}**")
    for t in TIERS:
        w(f"- {t}: {int((final['سطح (مشاور)'] == t).sum())}")
    w(f"- نسخهٔ اکسل با همهٔ ستون‌ها: `{Path(path).with_suffix('.xlsx').name}`\n")
    w("## هشدارها\n")
    for x in warns:
        w(f"- {x}")
    w("- کدها را پیش از ثبت با دفترچهٔ رسمی و اطلاعیه‌های اصلاحی سنجش تطبیق دهید.\n")
    w("## لیست\n")
    w("| # | کد | رشته | دانشگاه / محل تحصیل | دوره | شروع | سطح | نکته |")
    w("|---|---|---|---|---|---|---|---|")
    for _, r in final.iterrows():
        notes = [f for f in ["تعهد خدمت", "شهریه‌ای", "مصاحبه/شرایط خاص", "فقط بومی", "ظرفیت کم"] if r[f]]
        w(f"| {r['اولویت']} | {r['کد رشته محل']} | {r['عنوان رشته']} | {r['دانشگاه / مؤسسه']} | "
          f"{r['دوره تحصیلی']} | {r['شروع']} | {r['سطح (مشاور)']} | {'، '.join(notes)} |")
    w("\n## کدها برای ورود سریع (به ترتیب)\n")
    w("```")
    codes = final["کد رشته محل"].tolist()
    for i in range(0, len(codes), 10):
        w("  ".join(f"{i + j + 1:>3}) {c}" for j, c in enumerate(codes[i:i + 10])))
    w("```\n")
    w("## مراحل فیلتر\n")
    w("| مرحله | ردیف باقی‌مانده |\n|---|---|")
    for t, n in steps:
        w(f"| {t} | {n:,} |")
    Path(path).write_text("\n".join(L) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
