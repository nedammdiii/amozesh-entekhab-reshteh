# -*- coding: utf-8 -*-
"""Write the deliverable workbook: data sheet + choice list + summary + legend.

  python make_xlsx.py --final final.json --out book.xlsx --title "..." \
                      [--group "علوم تجربی" --year 1405] [--max-choices 150]
                      [--no-choice-list]

The summary sheet uses COUNTIF/SUMIF so it recalculates when the data changes.
openpyxl writes formulas without cached values, so they read back as None until
Excel opens the file - that is expected, not a bug.

The choice list (لیست ۱۵۰ انتخاب) is a dynamic list: the counsellor types a
priority number in the yellow «اولویت انتخاب» column of the data sheet and the
list sheet sorts those rows with SMALL + INDEX/MATCH. See
../references/choice-list.md before changing any of its formulas.
"""
import argparse, sys, io, json, collections, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

_ap = argparse.ArgumentParser()
_ap.add_argument("--final", required=True)
_ap.add_argument("--out", required=True)
_ap.add_argument("--title", default="دفترچه انتخاب رشته - آزمون سراسری")
_ap.add_argument("--group", help="گروه آزمایشی for the list title (default: read from --title)")
_ap.add_argument("--year", help="exam year for the list title (default: read from --title)")
_ap.add_argument("--max-choices", type=int, default=150,
                 help="how many choices the form allows (150 since 1403)")
_ap.add_argument("--no-choice-list", action="store_true",
                 help="plain workbook: no priority column, no list sheet")
ARGS = _ap.parse_args()
OUT = ARGS.out
CHOICE_LIST = not ARGS.no_choice_list
MAX_CHOICES = ARGS.max_choices

rows = json.load(open(ARGS.final, encoding="utf-8"))

_FA = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
_EN = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def fa(n):
    return str(n).translate(_FA)


def title_parts():
    """(group, year) from --group/--year, else from a title like
    «دفترچه انتخاب رشته گروه علوم تجربی - آزمون سراسری ۱۴۰۵»."""
    group, year = ARGS.group, ARGS.year
    if not group:
        m = re.search(r"گروه\s+(.+?)(?:\s+[-–—]|\s+آزمون|\s+\d|\s+[۰-۹]|$)", ARGS.title)
        group = m.group(1).strip() if m else ""
    if not year:
        m = re.findall(r"[0-9۰-۹٠-٩]{4}", ARGS.title)
        year = m[-1] if m else ""
    return group, fa(str(year).translate(_EN)) if year else ""


PRIORITY = "اولویت انتخاب"   # input column A — the counsellor types here
KEY = "کلید کمکی"            # hidden formula column, last in the table

DATA_COLS = [
    ("کد رشته محل", 13),
    ("عنوان رشته", 42),
    ("نحوه پذیرش", 20),
    ("دوره تحصیلی", 16),
    ("جنس پذیرش", 12),
    ("استان", 18),
    ("دانشگاه / مؤسسه", 52),
    ("نوع دانشگاه", 30),
    ("زیرمجموعه", 26),
    ("ظرفیت نیمسال اول", 10),
    ("ظرفیت نیمسال دوم", 10),
    ("ظرفیت کل", 9),
    ("ظرفیت زن", 9),
    ("ظرفیت مرد", 9),
    ("شرط بومی / سهمیه", 46),
    ("توضیحات", 60),
    ("بخش دفترچه", 40),
    ("منبع دوره تحصیلی", 18),
    ("صفحه PDF", 9),
]
# The priority column is built as column A from the start, so every summary
# formula below is written against the final column letters. Inserting it into
# a finished workbook instead shifts nothing (openpyxl's insert_cols moves no
# widths, table refs or formulas) — that is how the first version broke.
COLS = ([(PRIORITY, 11)] if CHOICE_LIST else []) + DATA_COLS + ([(KEY, 12)] if CHOICE_LIST else [])
NUMERIC = {"ظرفیت نیمسال اول", "ظرفیت نیمسال دوم", "ظرفیت کل", "ظرفیت زن", "ظرفیت مرد", "صفحه PDF"}

FONT = "Tahoma"
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT, size=10)
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YELLOW = PatternFill("solid", fgColor="FFF2CC")
PRIO_FONT = Font(name=FONT, size=10, bold=True, color="0000FF")
CENTER = Alignment(horizontal="center", vertical="center")

wb = Workbook()

# ------------------------------------------------------------------ data ----
ws = wb.active
ws.title = "رشته محل ها"
ws.sheet_view.rightToLeft = True

ws.append([c[0] for c in COLS])
for i, (name, width) in enumerate(COLS, start=1):
    cell = ws.cell(row=1, column=i)
    cell.font = HDR_FONT
    cell.fill = HDR_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions[get_column_letter(i)].width = width
ws.row_dimensions[1].height = 34

last = len(rows) + 1
COL_LETTER = {name: get_column_letter(i) for i, (name, _) in enumerate(COLS, start=1)}
# ROW()/SCALE breaks ties between equal priorities so MATCH never returns the
# same row twice. It must stay far below the smallest step a counsellor types
# (12.5, 12.25): with SCALE = 10^(digits of the last row + 3) it is < 0.001.
SCALE = 10 ** (len(str(last)) + 3)

for i, r in enumerate(rows, start=2):
    data = [r[c[0]] for c in DATA_COLS]
    if CHOICE_LIST:
        data = [None] + data + [f'=IF(ISNUMBER(A{i}),A{i}+ROW()/{SCALE},"")']
    ws.append(data)

for row in ws.iter_rows(min_row=2, max_row=last, max_col=len(COLS)):
    for cell in row:
        cell.font = BODY_FONT
        cell.border = BORDER
        header = COLS[cell.column - 1][0]
        if header == PRIORITY:
            cell.fill = YELLOW
            cell.font = PRIO_FONT
            cell.alignment = CENTER
        elif header in NUMERIC:
            cell.alignment = CENTER
        elif header in ("نحوه پذیرش", "دوره تحصیلی", "جنس پذیرش", "استان", "کد رشته محل", KEY):
            cell.alignment = CENTER
        else:
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=False)

tbl = Table(displayName="Reshtehha", ref=f"A1:{get_column_letter(len(COLS))}{last}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
ws.freeze_panes = "A2"

DATA = "'رشته محل ها'"


def rng(name):
    letter = COL_LETTER[name]
    return f"{DATA}!${letter}$2:${letter}${last}"


if CHOICE_LIST:
    P = COL_LETTER[PRIORITY]
    ws.column_dimensions[COL_LETTER[KEY]].hidden = True
    ws.freeze_panes = "D2"          # اولویت + کد + عنوان stay visible
    prio_cells = f"{P}2:{P}{last}"

    # A number typed on a Persian keyboard («۱۲») is TEXT to Excel, so ISNUMBER
    # ignores it and the row silently never reaches the list. Reject it at the
    # keyboard, paint any that slip in (paste) orange, and count them on the list.
    dv = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True,
                        showErrorMessage=True, errorStyle="stop",
                        errorTitle="اولویت نامعتبر",
                        error="فقط عدد مثبت با ارقام انگلیسی وارد کنید (مثلاً 12 یا 12.5). "
                              "اگر کیبورد فارسی است، عدد را با ارقام انگلیسی تایپ کنید.")
    ws.add_data_validation(dv)
    dv.add(prio_cells)
    ws.conditional_formatting.add(prio_cells, FormulaRule(
        formula=[f"AND(ISNUMBER({P}2),COUNTIF(${P}$2:${P}${last},{P}2)>1)"],
        fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(prio_cells, FormulaRule(
        formula=[f'AND({P}2<>"",NOT(ISNUMBER({P}2)))'],
        fill=PatternFill("solid", fgColor="F4B084"), font=Font(color="833C0B", bold=True)))

# --------------------------------------------------------- choice list ----
LIST_SHEET = f"لیست {fa(MAX_CHOICES)} انتخاب"
if CHOICE_LIST:
    ls = wb.create_sheet(LIST_SHEET, 1)
    ls.sheet_view.rightToLeft = True
    ls.sheet_properties.tabColor = "70AD47"
    KEYR, PRIOR = rng(KEY), rng(PRIORITY)
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right_wrap = Alignment(horizontal="right", vertical="center", wrap_text=True)

    group, year = title_parts()
    heading = f"لیست {fa(MAX_CHOICES)} انتخاب رشته"
    if group:
        heading += f" — گروه {group}"
    if year:
        heading += f" {year}"
    ls["A1"] = heading
    ls["A1"].font = Font(name=FONT, size=14, bold=True, color="1F4E79")
    ls.merge_cells("A1:L1")
    ls.row_dimensions[1].height = 24

    ls["A2"] = ("این برگه خودکار پر می‌شود؛ چیزی در آن تایپ نکنید. در برگه «رشته محل ها» کنار هر ردیف "
                "دلخواه، در ستون زرد «اولویت انتخاب» یک عدد بنویسید (۱، ۲، ۳ ...) — با ارقام انگلیسی. "
                "برای جا دادن بین دو ردیف از عدد اعشاری استفاده کنید (مثلاً 12.5 بین 12 و 13). "
                "برای حذف، عدد را پاک کنید.")
    ls["A2"].font = Font(name=FONT, size=9, color="595959")
    ls["A2"].alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
    ls.merge_cells("A2:L2")
    ls.row_dimensions[2].height = 32

    ls["A3"] = "تعداد انتخاب‌شده:"
    ls["C3"] = f"=COUNT({PRIOR})"
    ls["D3"] = "باقیمانده:"
    ls["E3"] = f"={MAX_CHOICES}-C3"
    ls["F3"] = (f'=TRIM(IF(C3>{MAX_CHOICES},"⚠ بیش از {fa(MAX_CHOICES)} ردیف انتخاب شده؛ '
                f'فقط {fa(MAX_CHOICES)} تای اول نمایش داده می‌شود. ","")'
                f'&IF(COUNTA({PRIOR})>C3,"⚠ "&(COUNTA({PRIOR})-C3)&" خانه در ستون اولویت عدد نیست '
                f'و در لیست نمی‌آید (نارنجی شده‌اند)؛ آن‌ها را با ارقام انگلیسی دوباره تایپ کنید.",""))')
    ls.merge_cells("A3:B3")
    ls.merge_cells("F3:L3")
    for a in ("A3", "D3"):
        ls[a].font = Font(name=FONT, size=10, bold=True)
        ls[a].alignment = right_wrap
    for a in ("C3", "E3"):
        ls[a].font = Font(name=FONT, size=11, bold=True, color="1F4E79")
        ls[a].alignment = center_wrap
    ls["F3"].font = Font(name=FONT, size=10, bold=True, color="C00000")
    ls["F3"].alignment = right_wrap
    ls.row_dimensions[3].height = 30

    ls["A4"] = ("نکته: اگر متن توضیحات کامل دیده نمی‌شود، ردیف‌های لیست را انتخاب کنید و "
                "Home ← Format ← AutoFit Row Height را بزنید (یا روی مرز بین شماره‌ی دو ردیف دابل‌کلیک کنید).")
    ls["A4"].font = Font(name=FONT, size=9, bold=True, color="2E75B6")
    ls["A4"].alignment = Alignment(horizontal="right", vertical="center")
    ls.merge_cells("A4:L4")

    # (header, width, source column in the data sheet, numeric?)
    # Numeric sources keep their numbers; text sources get &"" because INDEX on
    # an empty cell returns 0, which printed «0» in every blank توضیحات.
    LIST_COLS = [
        ("ردیف", 6, None, True),
        ("اولویت واردشده", 10, PRIORITY, True),
        ("کد رشته محل", 11, "کد رشته محل", False),
        ("عنوان رشته", 34, "عنوان رشته", False),
        ("دانشگاه / مؤسسه", 42, "دانشگاه / مؤسسه", False),
        ("استان", 14, "استان", False),
        ("دوره تحصیلی", 14, "دوره تحصیلی", False),
        ("نحوه پذیرش", 14, "نحوه پذیرش", False),
        ("جنس پذیرش", 10, "جنس پذیرش", False),
        ("ظرفیت کل", 8, "ظرفیت کل", True),
        ("شرط بومی / سهمیه", 45, "شرط بومی / سهمیه", False),
        ("توضیحات", 60, "توضیحات", False),
        ("هشدار", 10, None, False),
        ("شماره ردیف (کمکی)", 8, None, True),
    ]
    WRAP_RIGHT = {"عنوان رشته", "دانشگاه / مؤسسه", "شرط بومی / سهمیه", "توضیحات"}
    SMALL_FONT = {"شرط بومی / سهمیه", "توضیحات"}
    L = {h: get_column_letter(j) for j, (h, *_) in enumerate(LIST_COLS, start=1)}
    PR, IX, WARN = L["اولویت واردشده"], L["شماره ردیف (کمکی)"], L["هشدار"]
    HR = 5
    for j, (h, w, _, _) in enumerate(LIST_COLS, start=1):
        c = ls.cell(row=HR, column=j, value=h)
        c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, center_wrap, BORDER
        ls.column_dimensions[get_column_letter(j)].width = w
    ls.row_dimensions[HR].height = 30
    ls.column_dimensions[IX].hidden = True

    zebra = PatternFill("solid", fgColor="F2F6FB")
    fonts = {h: Font(name=FONT, size=9 if h in SMALL_FONT else 10, bold=(h == "ردیف"),
                     color="C00000" if h == "هشدار" else "000000") for h, *_ in LIST_COLS}
    end = HR + MAX_CHOICES
    for k in range(1, MAX_CHOICES + 1):
        r = HR + k
        ls.cell(row=r, column=1, value=k)
        ls[f"{IX}{r}"] = f'=IF(COUNT({KEYR})>={k},MATCH(SMALL({KEYR},{k}),{KEYR},0),"")'
        for j, (h, _, src, numeric) in enumerate(LIST_COLS, start=1):
            if src:
                tail = "" if numeric else '&""'
                ls.cell(row=r, column=j, value=f'=IF(${IX}{r}="","",INDEX({rng(src)},${IX}{r}){tail})')
        # the list is sorted, so equal priorities sit next to each other: flag
        # BOTH rows of a tie, as the red highlight in the data sheet does
        near = [f"{PR}{r}={PR}{r + d}" for d in (-1, 1) if HR < r + d <= end]
        if near:
            ls[f"{WARN}{r}"] = f'=IF(${IX}{r}="","",IF(OR({",".join(near)}),"تکراری",""))'
        for j, (h, *_) in enumerate(LIST_COLS, start=1):
            c = ls.cell(row=r, column=j)
            c.border = BORDER
            c.font = fonts[h]
            c.alignment = right_wrap if h in WRAP_RIGHT else center_wrap
            if k % 2 == 0:
                c.fill = zebra
    ls.freeze_panes = f"D{HR + 1}"
    ls.print_title_rows = f"{HR}:{HR}"
    ls.print_area = f"A1:{WARN}{end}"
    ls.page_setup.orientation = "landscape"
    ls.page_setup.fitToWidth = 1
    ls.page_setup.fitToHeight = 0
    ls.sheet_properties.pageSetUpPr.fitToPage = True

# --------------------------------------------------------------- summary ----
sm = wb.create_sheet("خلاصه آماری")
sm.sheet_view.rightToLeft = True
sm.column_dimensions["A"].width = 46
sm.column_dimensions["B"].width = 16
sm.column_dimensions["C"].width = 18


EMPTY_LABEL = "(خالی — در دفترچه درج نشده)"


def block(title, field, values, row):
    c = sm.cell(row=row, column=1, value=title)
    c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    c.fill = HDR_FILL
    for col, head in ((2, "تعداد کدرشته"), (3, "مجموع ظرفیت")):
        h = sm.cell(row=row, column=col, value=head)
        h.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
        h.fill = HDR_FILL
        h.alignment = Alignment(horizontal="center")
    col_rng, caprng = rng(field), rng("ظرفیت کل")
    r = row + 1
    for v in values:
        # An empty value (تجربی ۱۴۰۳: 73 شهید رجایی rows have no نحوه پذیرش) must
        # not be written as an empty label: COUNTIF against a blank cell counts 0,
        # and the block total silently came out 73 short. Name it and match "".
        sm.cell(row=r, column=1, value=v or EMPTY_LABEL).font = BODY_FONT
        crit = f"A{r}" if v else '""'
        f1 = sm.cell(row=r, column=2, value=f'=COUNTIF({col_rng},{crit})')
        f2 = sm.cell(row=r, column=3, value=f'=SUMIF({col_rng},{crit},{caprng})')
        for f in (f1, f2):
            f.font = BODY_FONT
            f.alignment = Alignment(horizontal="center")
        r += 1
    t = sm.cell(row=r, column=1, value="جمع")
    t.font = Font(name=FONT, size=10, bold=True)
    for col in (2, 3):
        s = sm.cell(row=r, column=col,
                    value=f"=SUM({get_column_letter(col)}{row+1}:{get_column_letter(col)}{r-1})")
        s.font = Font(name=FONT, size=10, bold=True)
        s.alignment = Alignment(horizontal="center")
    return r + 2


def order(field):
    c = collections.Counter(r[field] for r in rows)
    return [k for k, _ in c.most_common()]


r = 1
title = sm.cell(row=r, column=1, value="خلاصه آماری " + ARGS.title)
title.font = Font(name=FONT, size=12, bold=True)
r += 2
for label, field in (("نحوه پذیرش", "نحوه پذیرش"),
                     ("نوع دانشگاه", "نوع دانشگاه"),
                     ("دوره تحصیلی", "دوره تحصیلی"),
                     ("جنس پذیرش", "جنس پذیرش"),
                     ("زیرمجموعه (وزارتخانه)", "زیرمجموعه"),
                     ("استان", "استان")):
    r = block(label, field, order(field), r)

# ---------------------------------------------------------------- legend ----
lg = wb.create_sheet("راهنما")
lg.sheet_view.rightToLeft = True
lg.column_dimensions["A"].width = 26
lg.column_dimensions["B"].width = 110

def _uniq(col):
    """Distinct values of a column, most common first — keeps the legend honest."""
    return [k for k, _ in collections.Counter(r[col] for r in rows).most_common() if k]


CHOICE_NOTES = [
    ("اولویت انتخاب", "ستون ورودی زرد (ستون " + COL_LETTER.get(PRIORITY, "") + "): به هر ردیفی که می‌خواهید "
                      "در لیست باشد یک عدد اولویت بدهید، با ارقام انگلیسی. ترتیب لیست بر اساس کوچکی عدد است؛ "
                      "عدد اعشاری (مثلاً 12.5) ردیف را بین 12 و 13 می‌گذارد. عدد تکراری قرمز و ورودی "
                      "غیرعددی (مثلاً عددی که با کیبورد فارسی تایپ شده) نارنجی می‌شود."),
    (f"لیست {fa(MAX_CHOICES)} تایی", f"در برگه «{LIST_SHEET}» خودکار و مرتب ساخته می‌شود و تعداد انتخاب‌شده "
                                     "و باقیمانده را هم نشان می‌دهد. ستون «" + KEY + "» (ستون "
                                     + COL_LETTER.get(KEY, "") + "، مخفی) فرمول لیست است و نباید پاک "
                                     "یا جابه‌جا شود. برای فهرست خالی، عددهای ستون اولویت را پاک کنید."),
] if CHOICE_LIST else []

NOTES = [
    ("منبع", ARGS.title + " — استخراج خودکار از فایل PDF دفترچه "
             f"(صفحات {min(r['صفحه PDF'] for r in rows)} تا {max(r['صفحه PDF'] for r in rows)} فایل PDF)"),
    ("تعداد ردیف", f"{len(rows)} کدرشته‌محل (هر ردیف = یک کدرشته‌محل یکتا)"),
    ("", ""),
    *CHOICE_NOTES,
    ("کد رشته محل", "کد پنج‌رقمی مندرج در ستون «کدرشته محل» دفترچه."),
    ("عنوان رشته", "دقیقاً مطابق ستون «عنوان رشته» دفترچه."),
    ("نحوه پذیرش", "«با آزمون» یا «صرفا با سوابق تحصیلی» - مطابق ستون دفترچه."),
    ("دوره تحصیلی", "مقادیر موجود در این فایل: " + "، ".join(_uniq("دوره تحصیلی")) + ". "
                    "ستون «منبع دوره تحصیلی» نشان می‌دهد این مقدار مستقیم از ستون دفترچه آمده یا از عنوان بخش دفترچه استنتاج شده است."),
    ("جنس پذیرش", "«زن و مرد» / «فقط زن» / «فقط مرد» - بر اساس دو ستون زن و مرد در دفترچه "
                  "(خط تیره یعنی آن جنسیت پذیرش ندارد)."),
    ("استان", "استان محل تحصیل دانشگاه. برای جدول‌های سهمیه‌ای که در آن‌ها استان بالای جدول نیامده، "
              "استان از روی نام دانشگاه و جدول‌های اصلی همان دانشگاه استخراج شده است."),
    ("دانشگاه / مؤسسه", "نام کامل دانشگاه/دانشکده/مؤسسه به همراه محل تحصیل (در صورت درج در دفترچه)."),
    ("نوع دانشگاه", "مقادیر موجود در این فایل: " + "، ".join(_uniq("نوع دانشگاه")) + "."),
    ("زیرمجموعه", "وزارتخانهٔ متولی: " + "، ".join(_uniq("زیرمجموعه")) + "."),
    ("ظرفیت نیمسال اول/دوم", "عدد ستون «ظرفیت پذیرش نیمسال». خط تیره در دفترچه = صفر."),
    ("ظرفیت کل", "جمع ظرفیت نیمسال اول و دوم."),
    ("ظرفیت زن / ظرفیت مرد", "فقط برای ردیف‌هایی که دفترچه ظرفیت را به تفکیک جنسیت عدد زده است؛ "
                             "در بقیه ردیف‌ها خالی است."),
    ("شرط بومی / سهمیه", "برای جدول‌های سهمیه‌ای (بومی، مناطق محروم، بلایای طبیعی، مصوبه افزایش ظرفیت) "
                         "متن عنوان جدول در دفترچه."),
    ("توضیحات", "ستون «توضیحات» دفترچه (مصاحبه، خوابگاه، تعهد خدمت، بورس و ...)."),
    ("بخش دفترچه", "بخشی از دفترچه که ردیف از آن استخراج شده است."),
    ("صفحه PDF", "شماره صفحه در فایل PDF (نه شماره چاپ‌شده روی صفحه)."),
    ("", ""),
    ("نکته", "استخراج به‌صورت خودکار و بر اساس خطوط جدول و مختصات دقیق هر سلول در PDF انجام شده است؛ "
             f"هر {len(rows)} کد موجود در صفحات جدول‌ها استخراج و با شمارش کدهای هر صفحه "
             "صحت‌سنجی شده است (هیچ کدی جا نیفتاده)."),
]
hdr = lg.cell(row=1, column=1, value="راهنمای ستون‌ها")
hdr.font = Font(name=FONT, size=12, bold=True, color="FFFFFF")
hdr.fill = HDR_FILL
lg.cell(row=1, column=2).fill = HDR_FILL
for i, (k, v) in enumerate(NOTES, start=2):
    a = lg.cell(row=i, column=1, value=k)
    b = lg.cell(row=i, column=2, value=v)
    a.font = Font(name=FONT, size=10, bold=True)
    b.font = BODY_FONT
    a.alignment = Alignment(horizontal="right", vertical="top")
    b.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)

# Excel computes every formula when the file opens; without this the summary
# and the list can show stale (empty) cached values in some viewers.
wb.calculation.fullCalcOnLoad = True
wb.active = 0
wb.save(OUT)
print("saved:", OUT, "rows:", len(rows),
      f"| choice list: {LIST_SHEET} ({MAX_CHOICES} rows)" if CHOICE_LIST else "| choice list: off")
