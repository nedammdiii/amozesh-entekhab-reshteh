# Anatomy of a سنجش انتخاب‌رشته booklet PDF

Everything here was established empirically on the ۱۴۰۴ booklets for
**علوم تجربی** (407 pp), **ریاضی و فنی** (327 pp) and **علوم انسانی** (297 pp).
Read this before changing anything in `scripts/pdfgrid.py` or
`scripts/extract_rows.py` — most of the code exists to work around a specific
defect documented below.

**The single most important lesson from testing three booklets: the same
publisher typesets the same table differently in each one.** Header words are
split at different points, sections appear in a different order, and the
غیرانتفاعی/پیام‌نور tables drop a column. Anything matched literally — a header
word, a page range, a wing boundary — will break on the next booklet. Match
structurally and let `validate.py` be the gate.

## Contents

1. [Table anatomy](#1-table-anatomy)
2. [Column layout variants](#2-column-layout-variants)
3. [Finding rows](#3-finding-rows)
4. [Finding columns](#4-finding-columns)
5. [Merged cells](#5-merged-cells)
6. [The three text-layer defects](#6-the-three-text-layer-defects)
7. [Span colours](#7-span-colours)
8. [Caption grammar](#8-caption-grammar)
9. [The سهمیه‌ای trap](#9-the-سهمیهای-trap)
10. [Cell value vocabularies](#10-cell-value-vocabularies)
11. [1404 regression baseline](#11-1404-regression-baseline)

---

## 1. Table anatomy

The booklet is born-digital (real text layer, real vector rules — no OCR
needed) and every رشته‌محل table is drawn with the same column order,
**right to left**:

```
نحوه پذیرش │ [دوره تحصیلی] │ کدرشته محل │ عنوان رشته │ ظرفیت پذیرش نیمسال │ جنس پذیرش │ توضیحات
                                                       اول │ دوم        زن │ مرد
```

Two header rows: the outer one spans `ظرفیت پذیرش نیمسال` and `جنس پذیرش`, the
inner one splits each into `اول│دوم` and `زن│مرد`. In RTL the *first* sub-column
is the *rightmost*: `اول` sits at a **higher x** than `دوم`, `زن` higher than `مرد`.

A page holds one or more tables. Each table is preceded by a dark-red caption
naming استان and دانشگاه, and each table restarts with its own header row —
including continuation tables, whose caption is prefixed `ادامه` (in black, so
it does not pollute the red caption text).

## 2. Column layout variants

**The column set is not constant.** `survey.py --layouts` reports the distinct
header signatures. In 1404 تجربی there were four real ones:

| Signature | Pages | Note |
|---|---|---|
| `نحوه دوره کدرشته عنوان اول دوم زن مرد توضیحات` | 39–124, 165–236 | the full layout |
| `نحوه کدرشته عنوان اول دوم زن مرد توضیحات` | 250–403 | **no دوره column** — غیرانتفاعی / پیام‌نور, the دوره is implied by the section |
| `نحوه کدرشته عنوان اول دوم زن مرد دانشگاه توضیحات` | 125–162, 237–249 | left column is `دانشگاه محل تحصیل / توضیحات` — see §9 |
| `گروه کدرشته عنوان اول دوم زن مرد دانشگاه توضیحات` | 163–164 | `گروه تحصیلی` is the booklet's own **typo** for `نحوه پذیرش`; the cells hold «با آزمون»/«صرفا با سوابق تحصیلی». `HEADER_KEYS["nahve"]` accepts both words. |

ریاضی adds a fifth: the same full layout but with the code header printed as
**`کد رشته محل`** (three words) instead of `کدرشته محل`.

`survey.py --layouts` will also report junk signatures like `کدرشته` alone on
prose pages. Those are harmless: `extract_rows.py` only accepts a header whose
mapping covers every role in `NEEDED`.

**Never assume a variant from the page number.** Map the roles from each
table's own header row.

### Header words are split arbitrarily — never match them literally

The generator breaks header words at different points in each booklet:

| Printed header | تجربی | ریاضی | انسانی |
|---|---|---|---|
| کدرشته محل | `کدرشته` `محل` | `کد` `رشته` `محل` | `کدرشته` `محل` |
| توضیحات | `توضیحات` | `توضیحات` | `ت` `وضیحات` |
| اول | `اول` | `اول` | `ا` `ول` |
| مرد | `مرد` | `مرد` | `مر` `د` |

Exact word matching silently drops **whole sections** — it cost 109 pages of
ریاضی and 9 tables of انسانی before this was caught. So `map_header` assigns
every header word to a column first, concatenates each column's words (line by
line, right to left, no spaces), and matches `HEADER_SUBSTR` against *that*:
`کد`+`رشته`+`محل` → `کدرشتهمحل` → contains `کدرشته`. ✓

The same applies to the header-band test itself: `is_header_band` accepts
`کدرشته` **or** `کد`, always together with `عنوان` (which no data row carries).

## 3. Finding rows

Row bands come from the PDF's horizontal rules, not from text `y` clustering.

**The test is: does the rule cross the کدرشته column?** — not "is it
full width". A full-width test looks right and fails silently: on p177
(دانشگاه اطلاعات و امنیت ملی امام باقر) a single vertically merged توضیحات
cell spans ten rows, so those ten row separators start at the
توضیحات│مرد boundary rather than the table's left edge, and a full-width test
collapses ten rows into one band holding ten codes.

`extract_rows.py` derives the کدرشته column x-range from the actual 5-digit
code words on the page, then keeps every rule y whose merged segments span it.
Consecutive y values become the bands.

All rules are drawn as thin **rectangles** (`"re"` items), not `"l"` lines;
`pdfgrid.page_rects` handles both.

## 4. Finding columns

Column boundaries are the vertical rules that span the band.

**The header band has no usable vertical grid** — its cell borders are cut in
half by the زن/مرد sub-header row, so no vertical rule spans the whole header.
Take the grid from the first data band underneath instead, then assign header
words to columns by their x-centre. This is why `extract_rows.py` does
`next((b for b in band_bounds[bi:] if len(b) >= 6), [])`.

## 5. Merged cells

For each column independently, walk outward from the row's band until you hit a
horizontal rule that spans **that column's** x-range; the cell's text is
everything inside the resulting y-range. This is what `cell()` does, and it is
what makes the ten p177 rows each carry the full shared توضیحات.

## 6. The three text-layer defects

All three are real and all three are silent — the output looks like Persian,
just subtly wrong. Fixes live in `pdfgrid.py`.

### a. The لا ligature

The font emits `لا` as **a zero-width `ا` (U+0627) followed by `ل` (U+0644)**.
Left alone, `اسلامی` extracts as `اسالمی`, `اطلاعات` as `اطالعات`, `سلامت` as
`سالمت`. 1,256 occurrences in the 1404 booklet.

Detect by the zero width, not by the letters: in `span_chars_fixed`, an `ا`
whose bbox has zero width and whose next char is `ل` becomes `ل` + `ا`.
Shadda and tanwin are also zero-width — they are dropped by `_NORM`, not
reordered.

### b. Bidi bracket displacement

Characters inside a span are in logical order **except** punctuation that the
generator moved. A caption stored as
`")دانشگاه علوم پزشکی تبریز (محل تحصیل شهرستان اهر"` draws the leading `)` at
x≈148 (the far left, correct visually) while the rest runs 336→152.
Splitting on spaces alone glues `)` onto `دانشگاه` and strands the word at the
end of the line.

Fix: build words by **physical adjacency**. Start a new word whenever the gap
`max(cur.x0 - prev.x1, prev.x0 - cur.x1)` exceeds `GAP = 1.0` pt, then order
words by descending x. Measured distribution justifies the threshold:
within-word gaps are 0.0 pt (95,082 of 95,300 samples); space-separated gaps
are 2.0 pt. The two-sided `max` keeps Latin digit runs (which advance
left-to-right inside RTL text) intact.

`tidy()` then repairs the cosmetic spacing: `( ` → `(`, ` )` → `)`, etc.

### c. Dropped ZWNJ — only partly recoverable

`رشته‌ها` is stored as two touching glyph runs with no separator, so naive
joining yields `رشتهها`. But the generator *also* splits words mid-word at the
same zero gap — `صرف|ا`, `تحصیل|ی`, `رو|زانه`, `ب|ا` all appear in انسانی.

**Geometry cannot tell the two apart.** Measured gaps at real ZWNJ boundaries
(`پیوست|ها`: −0.36 … 0.0 pt) and at mid-word splits (`صرف|ا`: −0.0,
`تحصیل|ی`: −0.1, `رو|زانه`: −0.04) overlap completely. An "insert ZWNJ whenever
both sides are letters" rule therefore produces `صرفا ب‌ا سوابق تحصی‌لی`, which
blew `نحوه پذیرش` up from 2 values to 20 — a corrupted filter column.

So `rtl_join` joins touching fragments directly and restores a ZWNJ only for a
closed set of plural enclitics (`ZWNJ_SUFFIX` = ها / های / هایی), tested
against the fragment's leading letter run so that `«پیوست`+`ها»` still works.
Compound names (`علی‌آباد`, `جندی‌شاپور`) consequently lose their ZWNJ. That is
the deliberate trade: a cosmetic loss in a free-text column beats corrupting a
column people filter on.

This matters less than it sounds downstream — `build_table.clean()` strips ZWNJ
from every value anyway, so the delivered workbook is identical either way.

### Normalisation

`_NORM` maps ي→ی (U+064A→U+06CC) and ك→ک (U+0643→U+06A9) — the booklet mixes
Arabic and Persian forms in the same sentence — and strips tanwin, shadda,
tatweel and a symbol-font bullet. Do not add ZWNJ to `_NORM`: it is inserted
during joining and `normalize()` runs afterwards.

## 7. Span colours

| Colour | Size | Meaning |
|---|---|---|
| `0x800000` | 10 | dark-red **table caption** — `استان X - دانشگاه Y` |
| `0x0070C0` | 13–15 | blue **section heading** — the sections.json boundaries |
| `0x0000FF` | 15 | the printed **page number**. Not a heading — filter it out |
| `0xCC3300` | 11–12 | body-text emphasis in the prose chapters. Ignore |
| `0x000000` | 8 | table body |

Multi-line captions: group red words whose `cy` gap is ≤ 18 pt. Distinct
captions on one page are always separated by a table, so the gap is far larger.
The `ادامه` continuation prefix is black, so it never enters the caption text.

## 8. Caption grammar

Four shapes, all seen in 1404 (`parse_caption` in `build_table.py`):

| Shape | استان | دانشگاه |
|---|---|---|
| `استان X - دانشگاه Y` | X | Y |
| `دانشگاه پیام نور استان X - مرکز Y` | X | the whole caption |
| `مخصوص داوطلبان بومی استان X` | — (see §9) | — (from the left column) |
| `پذیرش از تمام متقاضیان سراسر کشور، با اولویت متقاضیان بومی استان X` | — (see §9) | — (from the left column) |

Province spelling varies (`کهگیلویه و بویر احمد` vs `کهکیلویه و بویراحمد`);
`PROV_ALIASES` normalises them. Matching is anchored on the literal word
`استان` first, then falls back to a bare longest-province-name search.

## 9. The سهمیه‌ای trap

In the بومی / مناطق محروم / بلایای طبیعی / مصوبه‌افزایش‌ظرفیت sections, the
caption's استان is **the applicant's home province, not the university's**.
Code 37653 sits under `مخصوص داوطلبان بومی استان محروم ایلام` but is taught at
دانشگاه شهید چمران اهواز, in خوزستان. Copying the caption's province into an
`استان` column would be wrong for every row in those sections.

Handle it as `build_table.py` does, driven by `"kind": "بومی"` in sections.json:

- the **university** is the left column's text before the first ` - `
  (`split_inst_desc`); the remainder is the real توضیحات;
- the **province** comes from a lookup table built in pass 1 from the
  `"kind": "استانی"` sections, keyed on `norm_inst()` (parentheses and
  `(محل تحصیل …)` suffixes stripped), with a longest-contained-name fallback;
- the caption text itself goes to a separate `شرط بومی / سهمیه` column.

In تجربی ۱۴۰۴ this resolved all 13,397 rows with zero unknown provinces.

**Where the lookup cannot work:** a university that appears *only* in a سهمیه
section has no استانی table to be looked up from. In ریاضی ۱۴۰۴,
`دانشکده علوم پزشکی و خدمات بهداشتی درمانی چابهار` shows up solely in the
کاردانی مناطق محروم tables, leaving 9 rows unresolved. `validate.py` fails on
that by design; the fix is the `province_overrides` map in sections.json, not a
code change.

## 9b. Section order is not fixed

تجربی opens with وزارت بهداشت (pp. 39–164) and continues into وزارت علوم.
ریاضی does the opposite — وزارت علوم from p. 36 and وزارت بهداشت at the very
end (pp. 316–323). This is why `wing` is a **per-section** field in
sections.json rather than a single split page, and why `uni_type` reads the
section's wing instead of comparing page numbers.

## 10. Cell value vocabularies

- **نحوه پذیرش** — exactly two values: `با آزمون`, `صرفا با سوابق تحصیلی`.
  Anything else means the column mapping is wrong.
- **دوره تحصیلی** — `روزانه`, `نوبت دوم`, `شهریه پرداز`, `پردیس خودگردان`,
  `مجازی`, `مشترک`, `روزانه - غیردولتی`, and for دانشگاه آزاد rows
  `آزاد تمام وقت` / `خودگردان آزاد` (the booklet's own wording — not a bug).
- **ظرفیت اول / دوم** — a number, or `-` for none.
- **زن / مرد** — the *word* `زن`/`مرد` when that gender is accepted with a
  shared capacity, a **number** when the booklet splits the capacity by gender
  (e.g. code 31821: زن=1, مرد=1, نیمسال دوم=2), or `-` when not accepted.
  `build_table.py` derives `جنس پذیرش` ∈ {`زن و مرد`, `فقط زن`, `فقط مرد`} and
  keeps the numeric split in separate `ظرفیت زن` / `ظرفیت مرد` columns.

## 11. 1404 regression baseline

Smoke test after touching `pdfgrid.py` or `extract_rows.py`. All three were
verified end-to-end; تجربی and ریاضی also through `make_xlsx` + `verify_xlsx`.

| Metric | تجربی | ریاضی | انسانی |
|---|---|---|---|
| config | `assets/sections-1404-tajrobi.json` | `assets/sections-1404-ryazi.json` | — |
| pages | `--first 39 --last 403` | `--first 36 --last 323` | `--first 36 --last 293` |
| rows | 13,397 | 12,029 | 10,560 |
| duplicate codes | 0 | 0 | 0 |
| missing codes | 0 | 0 | 0 |
| header problems | 0 | 0 | 0 |
| rows with no استان | 0 | 0 (needs the چابهار override) | — |
| distinct دانشگاه | 960 | 862 | — |
| distinct عنوان رشته | 269 | 251 | — |
| نحوه پذیرش values | 2 | 2 | 2 |
| total ظرفیت کل | 365,668 | 357,782 | — |

تجربی detail: 4,345 با آزمون / 9,052 سوابق تحصیلی · 12,160 زن و مرد /
615 فقط زن / 622 فقط مرد · 31 provinces.

`rows.json` and `final.json` for تجربی must come back **byte-identical** to the
shipped 1404 deliverable.

---

## 12. The ۱۴۰۳ booklets — five more ways a booklet can differ

Established on the ۱۴۰۳ reprints for **تجربی** (415 pp), **ریاضی** (362 pp) and
**انسانی** (316 pp). Configs: `assets/sections-1403-*.json`. Every fix below is
a no-op on a booklet that does not have the defect, so ۱۴۰۴ is unaffected.

### a. The text layer is shaped glyphs, not letters

۱۴۰۳ stores every Persian letter as its **Arabic Presentation Form**
(U+FB50–FDFF, U+FE70–FEFF) — `ﻛﺪرﺷﺘﻪ`, not `کدرشته`. Every literal match misses
and the booklet reads as "no tables found". `pdfgrid._DESHAPE` rebuilds the
base letters from the Unicode compatibility decompositions; LAM-ALEF forms
decompose to two letters and are emitted sharing one bbox so the word-splitter
cannot break them apart. `deshape_text()` does the same for a plain
`get_text()` string — `survey.table_pages` needs it too.

### b. Brackets can be stored already mirrored

۱۴۰۴ stores `)` first and the RTL rebuild lands it correctly; ۱۴۰۳ stores the
mirrored glyph and comes out `شیراز)محل تحصیل آباده(`. `pdfgrid.unmirror()`
decides from the result — a well-formed string never closes a bracket it never
opened — and mirrors the whole string. Do not make this unconditional.

### c. A 5-digit number is not necessarily a کدرشته

توضیحات cells carry fees (`27000`) and rank caps (`رتبه کشوری مجاز (19000)`).
`min/max` over every 5-digit word stretches the کدرشته range across half the
page, no row rule covers it, and a whole table collapses into one band
(انسانی p61 lost 17 rows that way). And one range is not enough anyway: تجربی
p144 holds tables whose columns sit 27pt apart. `extract_rows.code_columns()`
clusters the code words by x and keeps a cluster only when vertical rules hug
it within 15pt — real code columns measure ~6pt, توضیحات numbers 30–50pt.
`validate.py` uses the same helper, or it reports fees as missing codes.

### d. Not every table repeats the header, and not every header is complete

* تجربی p144 (امام باقر) runs straight from its caption into data with **no
  header row**, on a shifted grid. When the carried-over grid finds no code but
  the band's own grid has exactly one, re-anchor on the band's grid.
* ریاضی pp.153–171 and انسانی pp.123–127 leave the **`عنوان رشته` header cell
  blank**, so `عنوان` cannot anchor `is_header_band`. Anchor on "no 5-digit
  code in the band" plus two header markers, and cap the band at 40 words —
  the prose paragraph between the two tables on تجربی p223 says «کدرشته‌محل»,
  «دوره» and «ظرفیت» in passing and otherwise passes for a header.
* The تربیت دبیر شهید رجایی tables (تجربی 239–242, ریاضی 172–191, انسانی
  140–142) are a fifth layout: `کدرشته │ عنوان │ ظرفیت │ جنس │ دانشگاه یا پردیس`
  — one capacity column, one gender column, **no نحوه پذیرش at all**, and the
  left column headed `دانشگاه`, not `توضیحات`. `resolve_layout()` handles the
  collapsed pairs; the missing نحوه comes from the section.

### e. «گروه تحصیلی» is not always the نحوه پذیرش typo

In ۱۴۰۴ that header sat over «با آزمون»/«صرفا با سوابق تحصیلی». In ۱۴۰۳'s
کاردانی بهداشت tables the same header sits over a genuine گروه آزمایشی column
holding «علوم تجربی». Only the VALUE can tell them apart, so `extract_rows`
drops a نحوه cell that says neither «آزمون» nor «سوابق» and the section's
`"nahve"` key supplies the real one (those headings state it themselves).

### Two new sections.json keys

`from_caption` — the ministry can change PART-WAY down a page (تجربی p133,
ریاضی p47: وزارت بهداشت ends and وزارت علوم begins between two tables). Give
both sections that page and the later one the caption of its first table.

`nahve` — the value for sections whose tables have no such column. Leave it
out to keep the cells empty rather than guess.

### Other ۱۴۰۳ facts

* The reprints carry a 45° `www-kanoon-ir` watermark, one 80pt line whose bbox
  covers half the page, so its words land inside real cells. `page_words` skips
  non-horizontal lines.
* تجربی prints codes 12028–12030 **twice** (p69 and p223) with identical
  values. `build_table` keeps one row per code when the repeat says the same
  thing; `validate` separates identical repeats (fine) from conflicting ones
  (still fatal).
* Codes are often set flush against Persian text (`ايمني كار35901`), so a
  `\b\d{5}\b` census misses them. Guard on neighbouring digits instead.

### 1403 regression baseline

| Metric | تجربی | ریاضی | انسانی |
|---|---|---|---|
| pages | `--first 48 --last 411` | `--first 44 --last 358` | `--first 44 --last 312` |
| raw rows | 13,666 | 13,120 | 11,503 |
| final rows | 13,663 | 13,120 | 11,503 |
| missing codes | 0 | 0 | 0 |
| header problems | 0 | 0 | 0 |
| rows with no استان | 0 | 0 | 0 |
| distinct دانشگاه | 1,068 | 972 | 871 |
| distinct عنوان رشته | 281 | 263 | 250 |
| total ظرفیت کل | 366,340 | 363,258 | 344,169 |
| empty نحوه پذیرش (شهید رجایی) | 73 | 485 | 41 |
