# -*- coding: utf-8 -*-
"""Completeness gate + value inventories for an extracted booklet.

Exits non-zero when a کدرشته printed on a page that produced rows is missing
from the output. Read the inventories too: word-salad in a low-cardinality
column means the column mapping is wrong, not that the booklet is odd.

  python validate.py --pdf X.pdf --rows rows.json --final final.json \
                     [--first 39 --last 403]
"""
import argparse, collections, io, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from pdfgrid import open_doc, page_words, page_rects
from extract_rows import code_columns

CODE_RE = re.compile(r"^\d{5}$")
LOW_CARD = ("نحوه پذیرش", "دوره تحصیلی", "جنس پذیرش", "نوع دانشگاه", "زیرمجموعه", "استان")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--rows", required=True)
    ap.add_argument("--final")
    ap.add_argument("--first", type=int, required=True)
    ap.add_argument("--last", type=int, required=True)
    a = ap.parse_args()

    rows = json.load(open(a.rows, encoding="utf-8"))
    by_page = collections.defaultdict(set)
    for r in rows:
        by_page[r["page"]].add(r["code"])

    doc = open_doc(a.pdf)
    missing_total, no_table, extra_total = 0, [], 0
    print("=== 1. completeness: every code on the page vs every code extracted ===")
    for pno in range(a.first, a.last + 1):
        words = page_words(doc[pno - 1])
        vert, _ = page_rects(doc[pno - 1])
        found = [w for w in words if CODE_RE.match(w["text"])]
        # Only 5-digit words standing in a real کدرشته column count. A fee or a
        # rank printed inside a توضیحات cell is not a missing رشته‌محل.
        cols = code_columns(found, vert)
        codes = {w["text"] for w in found
                 if any(lo <= w["cx"] <= hi for lo, hi in cols)}
        got = by_page.get(pno, set())
        if not got:
            if codes:
                no_table.append(pno)
            continue
        miss = sorted(codes - got)
        extra = sorted(got - codes)
        if miss or extra:
            missing_total += len(miss)
            extra_total += len(extra)
            print(f"  p{pno}: missing={miss} extra={extra}")
    print(f"  MISSING CODES: {missing_total}")
    print(f"  EXTRA CODES (in output, not on page): {extra_total}")
    if no_table:
        print(f"  pages with codes but no table detected: {no_table}")
        print("  -> render one and check it really is a reference list, not a missed table")

    # Split duplicates: a booklet sometimes prints the same کدرشته‌محل in two
    # sections with identical values (تجربی ۱۴۰۳ p69 and p223). That is a
    # property of the source, not an extraction fault. A code that comes back
    # with DIFFERENT values IS a fault and still fails the gate.
    by_code = collections.defaultdict(list)
    for r in rows:
        by_code[r['code']].append(r)
    key = lambda r: (r['title'], r['cap1'], r['cap2'], r['zan'], r['mard'])
    repeats = [c for c, rs in by_code.items()
               if len(rs) > 1 and len({key(r) for r in rs}) == 1]
    dups = [c for c, rs in by_code.items()
            if len(rs) > 1 and len({key(r) for r in rs}) > 1]
    print('')
    print('=== 2. duplicates ===')
    print(f'  identical repeats printed twice in the booklet: {len(repeats)} {repeats[:10]}')
    print(f'  CONFLICTING duplicate codes: {len(dups)} {dups[:10]}')

    print("\n=== 3. raw cell inventories (expect tiny, clean sets) ===")
    for f in ("nahve", "dore", "zan", "mard"):
        c = collections.Counter(r[f] for r in rows)
        print(f"\n  {f}: {len(c)} distinct")
        for k, v in c.most_common(12):
            print(f"     {v:6d}  {k!r}")
    bad = collections.Counter()
    for r in rows:
        for f in ("cap1", "cap2"):
            if r[f] not in ("-", "") and not r[f].isdigit():
                bad[r[f]] += 1
    print(f"\n  non-numeric capacity cells: {bad.most_common(10)}")
    print(f"  empty titles: {sum(1 for r in rows if not r['title'])}")
    print(f"  empty captions: {sum(1 for r in rows if not r['caption'])}")

    ok = missing_total == 0 and extra_total == 0 and not dups

    if a.final and os.path.exists(a.final):
        fin = json.load(open(a.final, encoding="utf-8"))
        print(f"\n=== 4. final table: {len(fin)} rows ===")
        for col in LOW_CARD:
            c = collections.Counter(r[col] for r in fin)
            print(f"\n  {col}: {len(c)} distinct")
            for k, v in c.most_common(40):
                print(f"     {v:6d}  {k!r}")
        noprov = sum(1 for r in fin if not r["استان"])
        print(f"\n  rows with no استان: {noprov}")
        print(f"  distinct دانشگاه / مؤسسه: {len({r['دانشگاه / مؤسسه'] for r in fin})}")
        print(f"  distinct عنوان رشته: {len({r['عنوان رشته'] for r in fin})}")
        print(f"  total ظرفیت کل: {sum(r['ظرفیت کل'] for r in fin)}")
        odd = [k for k in {r["دانشگاه / مؤسسه"] for r in fin}
               if re.search(r"(دانشگاه|دانشکده|موسسه|مؤسسه|مرکز|واحد)\s*$", k)]
        print(f"  institution names ending in an institution word "
              f"(word-order bug): {len(odd)} {odd[:5]}")
        if noprov or odd:
            ok = False

    print("\n" + ("PASS" if ok else "FAIL"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
