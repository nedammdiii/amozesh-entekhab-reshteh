# -*- coding: utf-8 -*-
"""Geometry + text helpers for سنجش انتخاب‌رشته booklets.

Everything here works around defects in the booklet PDF text layer.
Read ../references/booklet-layout.md before changing any of it.
"""
import re

import fitz


_NORM = {
    "\u064a": "\u06cc",  # ARABIC YEH -> FARSI YEH
    "\u0643": "\u06a9",  # ARABIC KAF -> KEHEH
    "\u064b": "",        # tanwin fath (decorative here)
    "\u0651": "",        # shadda
    "\u0640": "",        # tatweel
    "\uf06c": " ",       # symbol-font bullet
}


def normalize(s: str) -> str:
    s = "".join(_NORM.get(c, c) for c in s)
    return " ".join(s.split())


_PRES_RANGES = ((0xFB50, 0xFDFF), (0xFE70, 0xFEFF))


def _build_deshape():
    """Arabic Presentation Forms -> base letters, from Unicode decompositions.

    Some booklets (e.g. the ۱۴۰۳ set) store the text layer as *shaped glyphs*
    — ﻓ/ﻌ/ﻝ codepoints in U+FB50..FDFF and U+FE70..FEFF — instead of base
    letters, so every literal match (`کدرشته`, `عنوان`, province names) misses
    and the whole booklet reads as "no tables found". Unicode gives each such
    form a compatibility decomposition back to its base letter(s); LAM-ALEF
    forms decompose to two (ل + ا) in logical order.

    Booklets that already use base letters contain none of these codepoints,
    so this table is a no-op for them.
    """
    import unicodedata
    m = {}
    for lo, hi in _PRES_RANGES:
        for o in range(lo, hi + 1):
            ch = chr(o)
            dec = unicodedata.decomposition(ch)
            if not dec:
                continue
            parts = [q for q in dec.split() if not q.startswith("<")]
            if parts:
                m[ch] = "".join(chr(int(q, 16)) for q in parts)
    return m


_DESHAPE = _build_deshape()


def _deshape_chars(chars):
    """Expand shaped glyphs; a ligature yields two chars sharing one bbox.

    The shared bbox keeps them physically touching, so the word-splitting gap
    test in page_words() never breaks a ligature in half.
    """
    out = []
    for c in chars:
        base = _DESHAPE.get(c["c"])
        if base is None:
            out.append(c)
        else:
            for b in base:
                out.append({"c": b, "bbox": c["bbox"]})
    return out


def deshape_text(s: str) -> str:
    """De-shape a plain get_text() string. No-op on base-letter booklets."""
    return "".join(_DESHAPE.get(c, c) for c in s)


def span_chars_fixed(span):
    """Chars of a span with the لا ligature (emitted as zero-width ا + ل) repaired."""
    chars = _deshape_chars(span["chars"])
    out = []
    i, n = 0, len(chars)
    while i < n:
        c = chars[i]
        w = c["bbox"][2] - c["bbox"][0]
        if c["c"] == "\u0627" and abs(w) < 0.01 and i + 1 < n and chars[i + 1]["c"] == "\u0644":
            bb = chars[i + 1]["bbox"]
            out.append(("\u0644", bb))
            out.append(("\u0627", bb))
            i += 2
            continue
        out.append((c["c"], c["bbox"]))
        i += 1
    return out


GAP = 1.0  # pt: bigger horizontal discontinuity than this starts a new word

# Plural enclitics — the only fragments that reliably earn a ZWNJ back.
ZWNJ_SUFFIX = {"ها", "های", "هایی"}
_LETTERS = set("آأؤإئءابپتثجچحخدذرزژسشصضطظعغفقکگلمنهوی")


def _lead_letters(s: str) -> str:
    """Leading Arabic-letter run — the fragment may carry punctuation («پیوست|ها»)."""
    i = 0
    while i < len(s) and s[i] in _LETTERS:
        i += 1
    return s[:i]


def page_words(page):
    """Words (dicts) built from characters, split on spaces AND physical gaps.

    Splitting on the physical gap matters because the PDF sometimes places a
    bracket at the far end of a line while keeping it adjacent in the char
    stream (a bidi artifact).
    """
    words = []
    d = page.get_text("rawdict")
    for b in d["blocks"]:
        if b.get("type", 1) != 0:
            continue
        for line in b["lines"]:
            # Table text is always horizontal. A rotated line is decoration —
            # the reprint watermark is a single 80pt line at 45 degrees whose
            # bbox covers half the page, so its words land inside real cells.
            dx, dy = line.get("dir", (1.0, 0.0))
            if abs(dy) > 0.05 or dx < 0.95:
                continue
            for span in line["spans"]:
                color, size, font = span["color"], span["size"], span["font"]
                cur, boxes, prev = [], [], None
                for ch, bb in span_chars_fixed(span):
                    if ch.isspace():
                        if cur:
                            words.append(_mkword(cur, boxes, color, size, font))
                        cur, boxes, prev = [], [], None
                        continue
                    if prev is not None:
                        gap = max(bb[0] - prev[2], prev[0] - bb[2])
                        if gap > GAP:
                            if cur:
                                words.append(_mkword(cur, boxes, color, size, font))
                            cur, boxes = [], []
                    cur.append(ch)
                    boxes.append(bb)
                    prev = bb
                if cur:
                    words.append(_mkword(cur, boxes, color, size, font))
    return [w for w in words if w["text"]]


def _mkword(chars, boxes, color, size, font):
    txt = normalize("".join(chars))
    x0 = min(b[0] for b in boxes)
    y0 = min(b[1] for b in boxes)
    x1 = max(b[2] for b in boxes)
    y1 = max(b[3] for b in boxes)
    return {"text": txt, "x0": x0, "y0": y0, "x1": x1, "y1": y1,
            "cx": (x0 + x1) / 2, "cy": (y0 + y1) / 2,
            "color": color, "size": size, "font": font}


def group_lines(words, line_tol=3.5):
    if not words:
        return []
    ws = sorted(words, key=lambda w: w["cy"])
    lines, cur = [], [ws[0]]
    for w in ws[1:]:
        if abs(w["cy"] - cur[-1]["cy"]) <= line_tol:
            cur.append(w)
        else:
            lines.append(cur)
            cur = [w]
    lines.append(cur)
    for ln in lines:
        ln.sort(key=lambda w: -w["x0"])
    return lines


def rtl_join(words, line_tol=3.5):
    """Order words RTL within each line, lines top->bottom, then join."""
    out = []
    for ln in group_lines(words, line_tol):
        parts, prev = [], None
        for w in ln:
            if prev is not None and max(w["x0"] - prev["x1"], prev["x0"] - w["x1"]) <= GAP:
                # Physically touching fragments belong to one word. Some of
                # those breaks are a dropped ZWNJ (پیوست|ها) and some are the
                # generator splitting mid-word (صرف|ا, تحصیل|ی, رو|زانه).
                # Geometry cannot tell them apart — measured gaps overlap
                # completely — so only a closed set of enclitics gets a ZWNJ
                # back. Guessing wider corrupts low-cardinality filter columns.
                sep = "‌" if _lead_letters(w["text"]) in ZWNJ_SUFFIX else ""
                parts[-1] = parts[-1] + sep + w["text"]
            else:
                parts.append(w["text"])
            prev = w
        out.append(" ".join(parts))
    return tidy(normalize(" ".join(out)))


_MIRROR = str.maketrans("()[]{}«»", ")(][}{»«")


def unmirror(s: str) -> str:
    """Undo pre-mirrored bidi brackets.

    Booklets differ: some store «)» first and the RTL rebuild lands it
    correctly, others store the already-mirrored glyph and come out as
    «شیراز)محل تحصیل آباده(». Detect it from the result — a well-formed string
    never closes a bracket it has not opened — and mirror the whole string.
    """
    depth = 0
    for c in s:
        if c in "([{«":
            depth += 1
        elif c in ")]}»":
            depth -= 1
            if depth < 0:
                return s.translate(_MIRROR)
    return s


def tidy(s: str) -> str:
    s = unmirror(s)
    s = s.replace("( ", "(").replace(" )", ")")
    s = re.sub(r"(?<=[^\s(\[{«])\(", " (", s)   # «شیراز(محل» -> «شیراز (محل»
    s = s.replace("« ", "«").replace(" »", "»")
    s = s.replace(" ،", "،").replace(" .", ".")
    return " ".join(s.split())


def page_rects(page):
    """Thin rects/lines of the page as (vertical, horizontal) segment lists."""
    vert, hor = [], []
    for d in page.get_drawings():
        for item in d["items"]:
            if item[0] == "re":
                r = item[1]
                if r.width < 1.6 and r.height > 1.5:
                    vert.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
                elif r.height < 1.6 and r.width > 1.5:
                    hor.append(((r.y0 + r.y1) / 2, r.x0, r.x1))
            elif item[0] == "l":
                a, b2 = item[1], item[2]
                if abs(a.x - b2.x) < 1.6 and abs(a.y - b2.y) > 1.5:
                    vert.append(((a.x + b2.x) / 2, min(a.y, b2.y), max(a.y, b2.y)))
                elif abs(a.y - b2.y) < 1.6 and abs(a.x - b2.x) > 1.5:
                    hor.append(((a.y + b2.y) / 2, min(a.x, b2.x), max(a.x, b2.x)))
    return vert, hor


def open_doc(path):
    return fitz.open(path)
