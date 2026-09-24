---
name: sanjesh-booklet-to-excel
description: >
  Convert an Iranian سازمان سنجش انتخاب‌رشته booklet PDF (دفترچه راهنمای انتخاب
  رشته آزمون سراسری / کنکور) into a clean, filter-ready Persian Excel workbook —
  one row per کدرشته‌محل, with نحوه پذیرش, دوره تحصیلی, عنوان رشته, کد رشته,
  جنس پذیرش, استان, دانشگاه, نوع دانشگاه (پیام نور / غیرانتفاعی / آزاد / دولتی /
  علوم پزشکی), ظرفیت and توضیحات as filterable columns. Works for any گروه آزمایشی
  (تجربی، ریاضی، انسانی، هنر، زبان) and any year. Use this skill whenever the user
  mentions دفترچه انتخاب رشته, دفترچه سنجش, دفترچه کنکور, کدرشته‌محل, رشته‌محل,
  انتخاب رشته, sanjesh, entekhab reshte, or asks to turn a کنکور/admission booklet
  PDF into Excel — and also whenever they hand over a large Persian PDF full of
  university-admission tables and want it tabulated, even if they never say
  "دفترچه" or "سنجش" explicitly. Prefer this skill over ad-hoc PDF parsing for any
  such booklet: it already handles the Persian text-layer defects that make naive
  extraction silently wrong.
---

# سنجش انتخاب‌رشته booklet → Excel

Turns a 400-page Persian admission booklet into one clean spreadsheet row per
کدرشته‌محل, verified against the PDF so no code is dropped.

The booklet is born-digital (real text layer, real vector table rules), so no
OCR is involved. But its text layer has three Persian defects that make naive
extraction *silently* wrong, and its tables use several different column
layouts. The scripts here handle all of that; your job is the discovery pass
and the verification gates.

## Before you start

Check the two dependencies import; `pip install pymupdf openpyxl` only if not:

```bash
python -c "import fitz, openpyxl; print(fitz.__doc__)"
```

Work in a scratch directory. `SK` below is this skill's folder,
`W` your working directory.

## The workflow

### 1. Survey the booklet

```bash
python "$SK/scripts/survey.py" --pdf BOOK.pdf --census
python "$SK/scripts/survey.py" --pdf BOOK.pdf --headings
python "$SK/scripts/survey.py" --pdf BOOK.pdf --layouts
```

- `--census` gives the `--first`/`--last` table page range and flags *decoy*
  pages (the رشته/مقطع reference lists, which are full of digit runs but carry
  no کدرشته header) and gaps.
- `--headings` lists the blue section headings with their page numbers. **These
  are the section boundaries** and they are the one thing that genuinely
  changes from year to year.
- `--layouts` lists the distinct table-header signatures with page ranges, and
  tells you which sections lack a `دوره تحصیلی` column.

Copy `assets/sections-1404-tajrobi.json` (or `-ryazi`, whichever is closer) to
`$W/sections.json` and rewrite every section from the `--headings` output. For
each section set:

- `dore` — `null` when `--layouts` says that page range **has** a دوره column;
  otherwise the value implied by the heading (`"غیرانتفاعی"`, `"پیام نور"`, or
  `"روزانه"` for the سهمیه/تعهد sections, which the booklet's own preamble
  states are روزانه).
- `kind` — `"استانی"` when the red caption names the university and its
  province; `"بومی"` for the سهمیه sections, where the caption names the
  *applicant's* home province and the university lives in the left column.
- `wing` — the ministry. **Do not assume an order**: تجربی opens with
  وزارت بهداشت, ریاضی puts it in the last few pages.

**Show the user the section table you derived and get confirmation before
continuing.** Everything downstream depends on it.

### 2. Look at the pages

Render one page per layout variant `--layouts` reported, plus the first page of
each section, and **actually open the PNGs with the Read tool**:

```bash
python "$SK/scripts/survey.py" --pdf BOOK.pdf --render 40 165 250 306 --out-dir "$W"
```

Confirm the column roles against the picture. Do not skip this — it is what
catches a section whose tables have a different column set. `--dump N` prints a
page's raw text if you need it.

### 3. Extract

```bash
python "$SK/scripts/extract_rows.py" --pdf BOOK.pdf --out "$W/rows.json" \
       --first 39 --last 403
```

Expect `problems: 0` and `duplicate codes: 0`. A "N codes" problem means a row
band swallowed several rows — a merged-cell case the code did not handle;
render that page and investigate before going on.

### 4. Build the table

```bash
python "$SK/scripts/build_table.py" --rows "$W/rows.json" \
       --sections "$W/sections.json" --out "$W/final.json"
```

Expect `rows with no province: 0` and `unresolved institutions: 0`.

An unresolved institution is normally one that appears *only* in a سهمیه
section, so there is no استانی table to read its province from. Add it to
`"province_overrides"` in sections.json (institution name → province) and
re-run. Resolve it from the institution's own name — never guess.

### 5. Validate — this is a gate, not a report

```bash
python "$SK/scripts/validate.py" --pdf BOOK.pdf --rows "$W/rows.json" \
       --final "$W/final.json" --first 39 --last 403
```

**`MISSING CODES` must be 0 and the script must exit 0.** Do not proceed
otherwise.

Then *read* the inventories it prints. `نحوه پذیرش` must have exactly two
values; `دوره تحصیلی`, `جنس پذیرش` and `نوع دانشگاه` must be short, clean sets.
Word-salad or a suspiciously large distinct count in a low-cardinality column
means the **column mapping is wrong**, not that the booklet is odd. Pages
reported as "codes but no table detected" should match the decoy pages from
step 1 — render one to confirm.

### 6. Write and verify the workbook

```bash
python "$SK/scripts/make_xlsx.py" --final "$W/final.json" --out "OUT.xlsx" \
       --title "دفترچه انتخاب رشته گروه علوم تجربی - آزمون سراسری ۱۴۰۵"
python "$SK/scripts/verify_xlsx.py" --final "$W/final.json" --xlsx "OUT.xlsx"
```

`verify_xlsx.py` must report `cell mismatches: 0`. Writing a cell is not proof
it holds what you meant — this caught a real bug on the first booklet.

The workbook has three sheets: the data table (RTL, AutoFilter, frozen header),
`خلاصه آماری` (COUNTIF/SUMIF, so it recalculates), and `راهنما`. The summary
formulas have no cached values until Excel opens the file — that is expected.
LibreOffice is usually unavailable on this machine, so do not promise
pre-computed summary numbers.

### 7. Spot-check before delivering

Render at least three pages spread across different sections and compare them
against the extracted rows for those pages. Only then hand the file over.

## What to tell the user

Report, in Persian:

- row count and the **0 missing codes** result, naming the page range covered;
- the column list and which sheets exist;
- these two honest caveats, every time:
  - **`دوره تحصیلی` is inferred** for the sections whose tables have no such
    column. Say how many rows, and point at the `منبع دوره تحصیلی` column,
    which records `ستون دفترچه` vs `استنتاج از بخش دفترچه` per row.
  - **`استان` is the university's province.** In the سهمیه sections the
    booklet's own heading names the *applicant's* home province instead; that
    text is kept separately in `شرط بومی / سهمیه`.

## Reference

`references/booklet-layout.md` — table anatomy, the column-layout variants, how
rows and columns are found from the vector rules, merged cells, the three
text-layer defects (لا ligature, bidi bracket displacement, dropped ZWNJ), span
colours, caption grammar, the سهمیه‌ای trap, and the 1404 regression baseline.

**Read it before editing any script.** Almost every line in `pdfgrid.py` and
`extract_rows.py` exists to work around a specific documented defect; changes
that look like simplifications tend to reintroduce silent corruption.

`assets/sections-1404-tajrobi.json` and `assets/sections-1404-ryazi.json` are
complete worked configs, and the reference doc carries the verified numbers for
تجربی، ریاضی and انسانی ۱۴۰۴. Re-running the pipeline against those booklets
must reproduce them exactly — and تجربی must come back byte-identical. Use that
as the smoke test after touching any script.
