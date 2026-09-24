# -*- coding: utf-8 -*-
"""Extract every رشته‌محل row from a سنجش انتخاب‌رشته booklet into rows.json.

Reads the table geometry from the PDF's own vector rules — never from the raw
text order, which is unreliable for RTL. See ../references/booklet-layout.md.

  python extract_rows.py --pdf X.pdf --out rows.json [--first 39 --last 403]
"""
import argparse, collections, io, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from pdfgrid import open_doc, page_words, page_rects, rtl_join

RED = 0x800000
CODE_RE = re.compile(r"^\d{5}$")


def cluster(values, tol=1.6):
    out = []
    for v in sorted(values):
        if out and v - out[-1][-1] <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(g) / len(g) for g in out]


def merge_segments(segs, tol=1.0):
    out = []
    for a, b in sorted(segs):
        if out and a <= out[-1][1] + tol:
            out[-1][1] = max(out[-1][1], b)
        else:
            out.append([a, b])
    return out


def covers(segs, lo, hi):
    return any(s[0] <= lo and s[1] >= hi for s in segs)


CODE_COL_SLACK = 15.0   # pt between a code cluster and its own column rules


def code_columns(code_words, vert):
    """x-ranges of the کدرشته columns on a page -> [(lo, hi), ...].

    Two traps, both real:
      * توضیحات cells hold 5-digit numbers of their own (fees like «27000»,
        «رتبه کشوری مجاز (19000)»). Taking min/max over every 5-digit word
        stretches the range across half the page, so NO row rule covers it and
        whole tables collapse into one band.
      * a page can hold tables with DIFFERENT column geometry (تجربی p144: the
        امام باقر table sits 27pt to the right of the ones above it), so one
        range is not enough either.

    Cluster the code words by x, then keep a cluster only when vertical rules
    hug it on both sides — that is what makes it a table column rather than a
    number sitting inside a wide free-text cell. Measured: real code columns
    are within ~6pt of their rules, توضیحات numbers 30-50pt away.
    """
    groups = []
    for w in sorted(code_words, key=lambda w: w["cx"]):
        if groups and w["cx"] - groups[-1][-1]["cx"] <= 12:
            groups[-1].append(w)
        else:
            groups.append([w])
    vx = sorted({round(v[0], 1) for v in vert})
    out = []
    for g in groups:
        lo = min(w["x0"] for w in g) - 1
        hi = max(w["x1"] for w in g) + 1
        left = max([x for x in vx if x <= lo], default=None)
        right = min([x for x in vx if x >= hi], default=None)
        if left is None or right is None:
            continue
        if lo - left <= CODE_COL_SLACK and right - hi <= CODE_COL_SLACK:
            out.append((lo, hi))
    return out


# NEVER match header words exactly — the generator breaks them arbitrarily and
# differently in every booklet: تجربی prints «کدرشته محل» where ریاضی prints
# «کد رشته محل», and انسانی emits «توضیحات» as «ت»+«وضیحات», «اول» as «ا»+«ول»,
# «مرد» as «مر»+«د». Instead, join each column's header text (spaces removed,
# line by line, right to left) and match these substrings against it.
# «گروه تحصیلی» is the booklet's own typo for «نحوه پذیرش».
HEADER_SUBSTR = {
    "code": ("کدرشته",),
    "title": ("عنوانرشته", "عنوان"),
    "dore": ("دورهتحصیلی", "دوره"),
    "nahve": ("نحوهپذیرش", "نحوه", "گروهتحصیلی"),
    # «دانشگاه» is the last resort: the تربیت دبیر شهید رجایی tables head their
    # left column «دانشگاه یا پردیس محل تحصیل/ محل خدمت» with no «توضیحات».
    "desc": ("وضیحات", "دانشگاه"),
    "first": ("اول",),
    "second": ("دوم",),
    "zan": ("زن",),
    "mard": ("مرد",),
    # Single-column variants. In the normal layout «ظرفیت پذیرش نیمسال» and
    # «جنس پذیرش» are OUTER headers spanning two sub-columns, so these roles
    # also match there — only ever use them when the split pair is absent.
    "cap": ("ظرفیت",),
    "jens": ("جنس",),
}
# «نحوه پذیرش» is not universal: the شهید رجایی tables have no such column and
# take their value from the section (sections.json "nahve").
BASE_NEEDED = ("code", "title", "desc")
# «نحوه پذیرش» only ever says «با آزمون» or «صرفا با سوابق تحصیلی».
NAHVE_WORDS = ("آزمون", "سوابق")


def resolve_layout(found):
    """-> (single_cap, single_jens, missing roles).

    Two table variants collapse a split pair into one column:
      ظرفیت  instead of ظرفیت پذیرش نیمسال (اول│دوم)
      جنس    instead of جنس پذیرش (زن│مرد)
    """
    single_cap = "cap" in found and ("first" not in found or "second" not in found)
    single_jens = "jens" in found and ("zan" not in found or "mard" not in found)
    missing = [k for k in BASE_NEEDED if k not in found]
    if single_cap:
        found["first"] = found["cap"]
        found.pop("second", None)
    else:
        missing += [k for k in ("first", "second") if k not in found]
    if not single_jens:
        missing += [k for k in ("zan", "mard") if k not in found]
    return single_cap, single_jens, missing


HEADER_MARKERS = {"عنوان", "نحوه", "گروه", "دوره", "ظرفیت", "جنس"}


def is_header_band(words):
    """A data row always carries its 5-digit code; a header band never does.

    «عنوان» alone is not a safe anchor: in some sections (ریاضی pp.153-171,
    انسانی pp.123-127) the booklet leaves the header cell above the رشته column
    blank, so the word is simply not there.
    """
    if any(CODE_RE.match(w["text"]) for w in words):
        return False
    # A header is one to three short lines. The prose paragraph sitting between
    # two tables (تجربی p223, ریاضی p150, انسانی p120) says «کدرشته‌محل»,
    # «دوره» and «ظرفیت» in passing, so vocabulary alone cannot reject it.
    if len(words) > 40:
        return False
    t = {w["text"] for w in words}
    # The same split-word defect map_header works around reaches this test too:
    # انسانی p194 stores the header cell as «کدرشت»+«ه» and p262 as
    # «کدرش»+«ته», so exact word matching rejects the band and drops the whole
    # page — silently, because a page with no table at all is not a missing
    # code. Match the band joined in reading order instead.
    flat = "".join(x["text"] for x in
                   sorted(words, key=lambda x: (round(x["y0"] / 6), -x["x0"])))
    # Two markers, not one: the prose paragraph between the two tables on
    # تجربی p223 says «کدرشته‌محل» and «دوره», and one marker would accept it.
    return ("کدرشته" in flat or "کد" in t) and len(t & HEADER_MARKERS) >= 2


def header_text(words, bounds):
    """Each column's header text: its words joined line by line, RTL, no spaces."""
    def col_of(w):
        for i in range(len(bounds) - 1):
            if bounds[i] - 1 <= w["cx"] <= bounds[i + 1] + 1:
                return i
        return None

    joined = collections.defaultdict(list)
    for w in words:
        c = col_of(w)
        if c is not None:
            joined[c].append(w)
    # line first, then right-to-left inside the line: «کد رشته» sits on the top
    # line and «محل» on the one below, so a pure x sort would interleave them.
    return {c: "".join(x["text"] for x in sorted(ws, key=lambda x: (round(x["y0"] / 6), -x["x0"])))
            for c, ws in joined.items()}


def map_header(words, bounds):
    text = header_text(words, bounds)
    found = {}
    for role, subs in HEADER_SUBSTR.items():
        for sub in subs:
            hit = [c for c in sorted(text) if sub in text[c]]
            if hit:
                found[role] = hit[0]
                break
    # Blank «عنوان رشته» header cell: in every observed layout the رشته column
    # is the one immediately left of کدرشته محل (RTL -> one index lower).
    if "title" not in found and "code" in found:
        cand = found["code"] - 1
        if cand >= 0 and cand not in found.values():
            found["title"] = cand
    return found


def extract(pdf, first, last):
    doc = open_doc(pdf)
    rows, problems = [], []
    # header text of columns that took no role -> pages. A new booklet can add
    # a column (a start-term or selection-type column, say); the rows still
    # extract, so nothing else would notice its data being dropped.
    unmapped = collections.defaultdict(list)
    last_caption = None

    for pno in range(first, last + 1):
        page = doc[pno - 1]
        words = page_words(page)
        vert, hor = page_rects(page)
        codes_on_page = [w for w in words if CODE_RE.match(w["text"])]
        if not vert or not codes_on_page:
            continue

        hy = collections.defaultdict(list)
        for y, x0, x1 in hor:
            hy[round(y, 1)].append((x0, x1))
        hy = {y: merge_segments(v) for y, v in hy.items()}

        # Row separators are the rules crossing the کدرشته column. Full-width is
        # the wrong test: a vertically merged توضیحات cell spans many rows.
        cols = code_columns(codes_on_page, vert)
        if not cols:
            continue
        seps = sorted(cluster([y for y, s in hy.items()
                               if any(covers(s, lo, hi) for lo, hi in cols)], 1.2))
        if len(seps) < 2:
            continue
        bands = list(zip(seps, seps[1:]))

        band_words = [[w for w in words if t + 1 <= w["cy"] <= b - 1] for t, b in bands]
        band_bounds = [cluster([v[0] for v in vert if v[1] <= t + 2 and v[2] >= b - 2])
                       for t, b in bands]

        reds = sorted([w for w in words if w["color"] == RED], key=lambda w: w["cy"])
        blocks, grp = [], []
        for w in reds:
            if grp and w["cy"] - grp[-1]["cy"] > 18:
                blocks.append(grp)
                grp = []
            grp.append(w)
        if grp:
            blocks.append(grp)
        captions = [(max(w["y1"] for w in g), rtl_join(g)) for g in blocks]

        def sep_covers(y, lo, hi):
            for yy, segs in hy.items():
                if abs(yy - y) <= 1.2 and covers(segs, lo + 2, hi - 2):
                    return True
            return False

        cur_map, cur_bounds = None, None
        cur_single_cap = cur_single_jens = False
        for bi, (top, bot) in enumerate(bands):
            band = band_words[bi]
            if not band:
                continue
            if is_header_band(band):                            # header band
                # Header cell borders are split by the زن/مرد sub-header row, so
                # take the column grid from the first data band underneath.
                bounds = next((b for b in band_bounds[bi:] if len(b) >= 6), [])
                cur_bounds = bounds if len(bounds) >= 6 else None
                cur_map = map_header(band, bounds) if cur_bounds else None
                cur_single_cap = cur_single_jens = False
                if cur_map is not None:
                    cur_single_cap, cur_single_jens, missing = resolve_layout(cur_map)
                    if missing:
                        problems.append(f"p{pno} header@{top:.0f} missing {missing}")
                        cur_map = None
                    else:
                        for c, t in header_text(band, bounds).items():
                            if t and c not in cur_map.values() and pno not in unmapped[t]:
                                unmapped[t].append(pno)
                continue
            if cur_map is None or cur_bounds is None:
                continue
            def codes_in(bounds):
                lo_c, hi_c = bounds[cur_map["code"]], bounds[cur_map["code"] + 1]
                return [w for w in band
                        if CODE_RE.match(w["text"]) and lo_c - 1 <= w["cx"] <= hi_c + 1]

            codes = codes_in(cur_bounds)
            if not codes and len(band_bounds[bi]) == len(cur_bounds):
                # A table can start WITHOUT repeating the header row (تجربی p144:
                # دانشگاه اطلاعات و امنیت ملی امام باقر follows its caption straight
                # into data) and its columns can sit at different x than the table
                # above. Same column count -> same roles; re-anchor on this grid.
                alt = codes_in(band_bounds[bi])
                if len(alt) == 1:
                    cur_bounds = band_bounds[bi]
                    codes = alt
            if len(codes) != 1:
                if codes:
                    problems.append(f"p{pno} band {top:.0f}-{bot:.0f}: {len(codes)} codes")
                continue

            def cell(role):
                """Text of a cell, following vertical merges above and below."""
                if role not in cur_map:
                    return ""
                i = cur_map[role]
                lo, hi = cur_bounds[i], cur_bounds[i + 1]
                t, b = top, bot
                j = bi
                while j > 0 and not sep_covers(bands[j][0], lo, hi):
                    j -= 1
                    t = bands[j][0]
                k = bi
                while k < len(bands) - 1 and not sep_covers(bands[k][1], lo, hi):
                    k += 1
                    b = bands[k][1]
                ws = [w for w in words
                      if t + 1 <= w["cy"] <= b - 1 and lo - 1 <= w["cx"] <= hi + 1]
                return rtl_join(ws)

            cap = ""
            for y, t in captions:
                if y <= top + 2:
                    cap = t
            if cap:
                last_caption = cap
            else:
                cap = last_caption or ""

            nahve = cell("nahve")
            if nahve and not any(k in nahve for k in NAHVE_WORDS):
                # «گروه تحصیلی» heads the نحوه پذیرش column in some booklets and a
                # genuine گروه آزمایشی column («علوم تجربی») in others. Only the
                # cell value can tell them apart; the section supplies the rest.
                nahve = ""

            if cur_single_jens:
                j = cell("jens")
                zan = "زن" if "زن" in j else "-"
                mard = "مرد" if "مرد" in j else "-"
            else:
                zan, mard = cell("zan"), cell("mard")

            rows.append({
                "page": pno,
                "code": codes[0]["text"],
                "title": cell("title"),
                "nahve": nahve,
                "dore": cell("dore"),
                "cap1": cell("first"),
                "cap2": "" if cur_single_cap else cell("second"),
                "zan": zan,
                "mard": mard,
                "desc": cell("desc"),
                "caption": cap,
            })
    return rows, problems, unmapped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--first", type=int)
    ap.add_argument("--last", type=int)
    a = ap.parse_args()

    doc = open_doc(a.pdf)
    first, last = a.first or 1, a.last or doc.page_count
    rows, problems, unmapped = extract(a.pdf, first, last)

    dups = [c for c, n in collections.Counter(r["code"] for r in rows).items() if n > 1]
    print(f"pages {first}-{last}")
    print(f"rows: {len(rows)}")
    print(f"distinct codes: {len({r['code'] for r in rows})}")
    print(f"duplicate codes: {len(dups)} {dups[:10]}")
    print(f"problems: {len(problems)}")
    for p in problems[:40]:
        print("  ", p)
    print(f"header columns with no role: {len(unmapped)}")
    for t, pages in sorted(unmapped.items(), key=lambda kv: kv[1][0]):
        print(f"   {t!r} on {len(pages)} pages, first p{pages[0]}")
    if unmapped:
        print("   -> a column the scripts do not know. Its data is NOT in rows.json.")
        print("      Render one of those pages and decide (references/booklet-layout.md §13).")
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    print("wrote", a.out)


if __name__ == "__main__":
    main()
