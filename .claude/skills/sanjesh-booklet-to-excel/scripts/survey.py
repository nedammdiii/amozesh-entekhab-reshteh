# -*- coding: utf-8 -*-
"""Discovery pass over a سنجش انتخاب‌رشته booklet.

Run this FIRST on any new booklet. It tells you which pages hold tables, what
the section headings are (they become sections.json) and which column layouts
the tables use. Nothing here writes data files — it only reports.

  python survey.py --pdf X.pdf --census
  python survey.py --pdf X.pdf --headings
  python survey.py --pdf X.pdf --layouts
  python survey.py --pdf X.pdf --render 40 165 250 --out-dir DIR
  python survey.py --pdf X.pdf --dump 154
"""
import argparse, collections, io, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from pdfgrid import open_doc, page_words, rtl_join, deshape_text

RED = 0x800000          # table caption: «استان X - دانشگاه Y»
BLUE = 0x0070C0         # section heading
PAGENO = 0x0000FF       # printed page number — never a heading
CODE_RE = re.compile(r"^\d{5}$")
# NOT a word-boundary match: the booklet often sets a code flush against the
# Persian text next to it («...ايمني كار35901 روزانه»), and a letter/digit
# junction is no word boundary. Guard on neighbouring DIGITS instead.
CODE_ANY = re.compile(r"(?<!\d)\d{5}(?!\d)")


def runs(pages):
    out = []
    for p in sorted(pages):
        if out and p == out[-1][1] + 1:
            out[-1][1] = p
        else:
            out.append([p, p])
    return [tuple(r) for r in out]


def table_pages(doc):
    """Pages holding a رشته‌محل table.

    Both signals are needed: the header words also appear in the intro prose,
    and the رشته/مقطع reference lists are full of digit runs but carry no
    header. Spaces are stripped before matching because the header is printed
    «کدرشته محل» in some booklets and «کد رشته محل» in others.
    """
    out = []
    for p in range(doc.page_count):
        t = deshape_text(doc[p].get_text())
        flat = re.sub(r"\s+", "", t.replace("ك", "ک").replace("ي", "ی"))
        # «عنوان رشته» is not always printed — some sections leave that header
        # cell blank — so accept any of the other header labels as the partner.
        partner = any(k in flat for k in ("عنوانرشته", "نحوهپذیرش", "جنسپذیرش"))
        if "کدرشته" in flat and partner and CODE_ANY.search(t):
            out.append(p + 1)
    return out


def census(doc):
    """Locate the رشته‌محل tables and flag pages that only look like tables."""
    hits = table_pages(doc)
    codes = {p + 1: len(CODE_ANY.findall(doc[p].get_text()))
             for p in range(doc.page_count)}
    print(f"pages: {doc.page_count}")
    print(f"table pages (carry the «کدرشته» header): {len(hits)}")
    print(f"page runs: {runs(hits)}")
    if not hits:
        print("\nNo رشته‌محل tables found — is this the right booklet?")
        return
    print(f"\nSUGGESTED --first {min(hits)} --last {max(hits)}")

    decoys = [p for p, n in codes.items()
              if n >= 8 and p not in hits and min(hits) - 5 <= p <= max(hits) + 5]
    if decoys:
        print(f"\nDECOY pages (many digit runs, no کدرشته header): {runs(decoys)}")
        print("These are usually the رشته/مقطع reference lists. Render one to confirm,")
        print("then expect validate.py to report them as 'no table detected'.")
    gaps = [p for p in range(min(hits), max(hits) + 1) if p not in hits]
    if gaps:
        print(f"\nGAP pages inside the table range: {runs(gaps)}")


def colored_lines(page, colors, minsize=0.0):
    """Lines of same-coloured text, reassembled right-to-left."""
    rows = collections.defaultdict(list)
    for w in page_words(page):
        if w["color"] in colors and w["size"] >= minsize and w["text"]:
            rows[round(w["cy"], 0)].append(w)
    out = []
    for k in sorted(rows):
        out.append((min(w["y0"] for w in rows[k]), rtl_join(rows[k])))
    return out


def headings(doc, first, last):
    """Blue section headings -> the skeleton of sections.json."""
    print("=== BLUE section headings (candidate section boundaries) ===")
    for pno in range(first, last + 1):
        for y, t in colored_lines(doc[pno - 1], {BLUE}, minsize=12.5):
            if t.strip():
                print(f"p{pno:>4}  y={y:6.1f}  {t}")
    print("\n=== RED table captions: first 3 pages and last 2, plus total ===")
    total, sample = 0, []
    for pno in range(first, last + 1):
        caps = [t for _, t in colored_lines(doc[pno - 1], {RED}) if t.strip()]
        total += len(caps)
        if pno < first + 3 or pno > last - 2:
            sample += [f"p{pno}: {c}" for c in caps]
    for s in sample[:25]:
        print(" ", s)
    print(f"\ntotal red caption lines: {total}")
    print("\nNote: 0x0000FF size-15 spans are the printed page number, not headings.")


def layouts(doc, first, last):
    """Distinct table-header signatures = the column layout variants."""
    sig = collections.Counter()
    pages = collections.defaultdict(list)
    for pno in range(first, last + 1):
        words = page_words(doc[pno - 1])
        for k in [w for w in words if w["text"] in ("کدرشته", "کد")]:
            blk = [w for w in words if k["y0"] - 12 <= w["y0"] <= k["y0"] + 22]
            if "عنوان" not in {w["text"] for w in blk}:
                continue
            blk.sort(key=lambda w: -w["x0"])
            roles = tuple(w["text"] for w in blk
                          if w["text"] in ("نحوه", "گروه", "دوره", "کدرشته", "کد", "عنوان",
                                           "اول", "دوم", "زن", "مرد", "توضیحات", "وضیحات",
                                           "دانشگاه"))
            sig[roles] += 1
            pages[roles].append(pno)
    print(f"=== {len(sig)} distinct header signatures ===")
    for roles, n in sig.most_common():
        print(f"\n{n:5d} tables | {' '.join(roles)}")
        print(f"        pages: {runs(pages[roles])}")
        print(f"        has دوره تحصیلی column: {'YES' if 'دوره' in roles else 'NO -> must be inferred per section'}")


def render(doc, pages, out_dir, dpi):
    os.makedirs(out_dir, exist_ok=True)
    for p in pages:
        f = os.path.join(out_dir, f"p{p}.png")
        doc[p - 1].get_pixmap(dpi=dpi).save(f)
        print(f)


def dump(doc, pages):
    for p in pages:
        print("=" * 30, "PAGE", p)
        print(deshape_text(doc[p - 1].get_text()))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--census", action="store_true")
    ap.add_argument("--headings", action="store_true")
    ap.add_argument("--layouts", action="store_true")
    ap.add_argument("--render", nargs="+", type=int)
    ap.add_argument("--dump", nargs="+", type=int)
    ap.add_argument("--out-dir", default=".")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--first", type=int)
    ap.add_argument("--last", type=int)
    a = ap.parse_args()

    doc = open_doc(a.pdf)
    first = a.first or 1
    last = a.last or doc.page_count
    if (a.headings or a.layouts) and not a.first:
        hits = table_pages(doc)
        if hits:
            first, last = min(hits), max(hits)

    if a.census:
        census(doc)
    if a.headings:
        headings(doc, first, last)
    if a.layouts:
        layouts(doc, first, last)
    if a.render:
        render(doc, a.render, a.out_dir, a.dpi)
    if a.dump:
        dump(doc, a.dump)
    if not any([a.census, a.headings, a.layouts, a.render, a.dump]):
        ap.error("pick at least one of --census --headings --layouts --render --dump")


if __name__ == "__main__":
    main()
