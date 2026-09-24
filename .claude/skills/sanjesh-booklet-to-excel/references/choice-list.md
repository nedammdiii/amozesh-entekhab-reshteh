# The dynamic 150-choice list (لیست ۱۵۰ انتخاب)

The user of the workbook is a student counsellor. For each candidate they pick
about 150 of the ~13,000 کدرشته‌محل rows and put them in order. Copying rows
one by one into another file was slow and error-prone, so the workbook does
it with formulas: the counsellor types a **priority number** next to each
row they want, and a separate sheet shows those rows sorted.

The technique is a *dynamic list* driven by a *helper column*, using
`SMALL` + `INDEX/MATCH`. Everything lives in `scripts/make_xlsx.py`.
`scripts/verify_xlsx.py` checks the wiring statically and
`scripts/test_choice_list.py` proves the results by computing them.

## 1. Data sheet `رشته محل ها`

| Column | Header | Content |
|---|---|---|
| A | `اولویت انتخاب` | input: yellow `FFF2CC`, bold blue, centred, **shipped empty** |
| B … T | the 19 data columns | unchanged, in the old order |
| U (last) | `کلید کمکی` | hidden: `=IF(ISNUMBER(A2),A2+ROW()/SCALE,"")` |

Both extra columns belong to the Excel Table `Reshtehha`, so the table can be
sorted and filtered with the priorities travelling with their rows. The
panes are frozen at `D2`, which keeps priority, code and title visible.

**The priority column is written as column A from the start.** Inserting it
into a finished workbook with openpyxl's `insert_cols` moves no column widths,
no table ref and no formula on other sheets. Every summary formula would then
read the wrong column. Because `COL_LETTER` is derived from `COLS`, every
formula in the file is written against the final letters.

### The helper key

`ROW()/SCALE` breaks ties. Two rows with the same priority still get
distinct keys, so `MATCH` never returns the same row twice.
`SCALE = 10^(digits of the last row + 3)`, which is `1e8` for a
13,000-row booklet, so the tie-break stays below 0.001. Priorities with up to
three decimals (12.5, 12.25, 12.125) keep their order.

### Input guards on column A

- **Data validation:** decimal > 0, with a stop message in Persian. A number
  typed on a Persian keyboard (`۱۲`) is **text** to Excel. `ISNUMBER` ignores
  it, so without the guard the row silently never reaches the list.
- **Conditional format, orange `F4B084`:** non-empty and not a number. This
  catches values that bypass validation by paste.
- **Conditional format, red `FFC7CE` / `9C0006`:** duplicate priority,
  `AND(ISNUMBER(A2),COUNTIF($A$2:$A$N,A2)>1)`.

## 2. List sheet `لیست ۱۵۰ انتخاب`

It is the second tab, RTL, with tab colour `70AD47`.

| Row | Content |
|---|---|
| 1 | `لیست ۱۵۰ انتخاب رشته — گروه {group} {year}` (from `--title`, or `--group`/`--year`) |
| 2 | one-paragraph instructions |
| 3 | `تعداد انتخاب‌شده` `=COUNT(prio)`, `باقیمانده` `=MAX-C3`, warnings |
| 4 | the AutoFit Row Height hint |
| 5 | headers (repeated on every printed page) |
| 6 … 5+MAX | the list |

Columns: ردیف · اولویت واردشده · کد رشته محل · عنوان رشته · دانشگاه / مؤسسه ·
استان · دوره تحصیلی · نحوه پذیرش · جنس پذیرش · ظرفیت کل · شرط بومی / سهمیه ·
توضیحات · هشدار · شماره ردیف (کمکی, hidden N).

Formulas for list row *k* on sheet row *r*, where `KEY` is the helper column
range:

```
N  (hidden row index)   =IF(COUNT(KEY)>=k, MATCH(SMALL(KEY,k),KEY,0), "")
text columns            =IF($Nr="","",INDEX(src,$Nr)&"")
numeric columns         =IF($Nr="","",INDEX(src,$Nr))        (priority, ظرفیت کل)
هشدار                   =IF($Nr="","",IF(OR(Br=B(r-1),Br=B(r+1)),"تکراری",""))
```

The row-3 warning cell is built from two warnings, joined with `TRIM`:

- more than MAX numbers entered: only the first MAX are shown;
- `COUNTA(prio) > COUNT(prio)`: some priority cells hold text (the Persian-digit
  case). The warning gives how many.

### Why each formula looks the way it does

1. **`&""` on text columns only.** `INDEX` on an empty cell returns `0`, and
   every blank توضیحات showed «0». Numeric columns must not get the suffix,
   or they turn into text.
2. **The duplicate flag checks both neighbours.** The list is sorted, so equal
   priorities are adjacent. Checking only the row above flagged just the
   second of a pair, while the data sheet paints both red. The first row
   omits the check against the header, and the last row omits the check
   against the empty row below, where Excel would compare with 0.
3. **No `FILTER`, `SORT`, `SORTBY`, `XLOOKUP` or `XMATCH`.** Excel 2016/2019,
   still common in Iran, lacks them, and LibreOffice cannot compute them for
   the test. Only `SMALL`, `INDEX`, `MATCH`, `COUNT`, `COUNTA`, `IF`, `OR`
   and `TRIM` are used.
4. **Row height.** توضیحات runs to ~280 characters and شرط بومی to ~240.
   Excel does not re-fit a row when a formula result changes, and there is no
   formula-only fix. The columns are wide (34 / 42 / 45 / 60) and wrapped,
   and row 4 tells the user to run AutoFit Row Height.
5. **`fullCalcOnLoad`** is set, so Excel computes every formula when the file
   opens. openpyxl stores no cached values.
6. **MAX is a parameter** (`--max-choices`, default 150). The sheet name, the
   title, the counter and the warning all follow it.

## 3. Formatting

Headers use `1F4E79` with white Tahoma 10 bold. Rows alternate with `F2F6FB`,
and every cell wraps. The شرط بومی and توضیحات columns use Tahoma 9. The panes
are frozen at `D6`. For printing, the sheet is landscape and fits 1 page wide,
with header row 5 repeated and the print area ending at the هشدار column.

## 4. What the tests prove

`verify_xlsx.py` runs without Excel or LibreOffice:

- The priority column is empty. A test copy must never be delivered.
- Every helper key points at its own row.
- The table ref covers all columns.
- Every `INDEX` on the list reads the data column whose header equals its own.
  The only mapping is `اولویت واردشده` → `اولویت انتخاب`.
- Every row-index formula reads the helper column.

A shifted letter shows the wrong field under every heading without a single
`#REF!`, so this check matters.

`test_choice_list.py` uses LibreOffice. It has to be `libreoffice-calc`: a
core-only install cannot open xlsx, and the script detects that. It runs two
scenarios on copies.

- **A:** sheet rows 10, 500, 40, 3000, 7000 get 1, 2, 2.5, 3, 3. A row with an
  empty توضیحات gets 4, and one cell gets `«۵»` as text. Expected results:
  - order 1, 2, 2.5, 3, 3, 4, with every listed field equal to its source cell;
  - «تکراری» on both 3s;
  - blank, not «0», for the empty توضیحات;
  - count 6 and remaining 144;
  - the text-entry warning;
  - every line of every summary block equal to final.json (count and
    capacity), including a «(خالی …)» line for empty values;
  - no error value anywhere.
- **B:** MAX+1 priorities. Expect count 151, remaining −1, the over-limit
  warning, and row 150 holding the 150th code.

Mutation-checked when written. A wrong column letter, a priority left in the
file and dropping `&""` each make the matching test fail. On the real تجربی
۱۴۰۳ booklet the summary check also caught the empty-label bug, where the
نحوه پذیرش block came out 73 short (see booklet-layout.md §11).

## 5. Downstream readers

The `konkur-entekhab-reshteh` skill's `booklet.py` still finds the data
sheet. It has 19 known headers against the list sheet's 10. It reports
`اولویت انتخاب` and `کلید کمکی` as unrecognised columns and ignores them.
Verified on the new workbook.
