# -*- coding: utf-8 -*-
"""Turn rows.json into the final, filter-friendly table (final.json).

Adds the derived columns: استان, دانشگاه, نوع دانشگاه, جنس پذیرش, ظرفیت,
and fills دوره تحصیلی for the sections whose tables carry no such column.

  python build_table.py --rows rows.json --sections sections.json --out final.json

sections.json shape (page ranges are the one thing that changes every year —
get them from `survey.py --headings`):

{
  "sections": [
    {"from": 39, "to": 124, "name": "...", "dore": null,
     "kind": "استانی", "wing": "وزارت بهداشت (علوم پزشکی)"},
    {"from": 250,"to": 305, "name": "...", "dore": "غیرانتفاعی",
     "kind": "استانی", "wing": "وزارت علوم، تحقیقات و فناوری"}
  ]
}

`dore`  : value to use when the table has no «دوره تحصیلی» column (null = the
          column exists, take it from the PDF).
`kind`  : "استانی"  -> the red caption names the university and its province.
          "بومی"    -> the caption names the APPLICANT's home province; the
                       university lives in the left column instead.
`wing`  : the ministry the section belongs to. Per-section, not a single split
          page: تجربی puts وزارت بهداشت first, ریاضی puts it last.
`from_caption` (optional): the caption of this section's first table, used when
          the previous section ends PART-WAY down the same page. Give both
          sections that page and the rows are split at the caption.
`nahve` (optional): value for «نحوه پذیرش» in sections whose tables carry no
          such column (the تربیت دبیر شهید رجایی ones). Leave it out to keep
          those cells empty rather than guess.

Optional top-level key:

  "province_overrides": {"دانشکده علوم پزشکی ... چابهار": "سیستان و بلوچستان"}

A university that appears ONLY in a سهمیه section has no استانی table to be
looked up from, so its province cannot be derived. validate.py fails on that;
add the institution here (the name as printed, or any prefix of it) and re-run.
"""
import argparse, collections, io, json, re, sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

ZWNJ = "‌"

PROVINCES = [
    "آذربایجان شرقی", "آذربایجان غربی", "اردبیل", "اصفهان", "البرز", "ایلام", "بوشهر",
    "تهران", "چهارمحال و بختیاری", "خراسان جنوبی", "خراسان رضوی", "خراسان شمالی",
    "خوزستان", "زنجان", "سمنان", "سیستان و بلوچستان", "فارس", "قزوین", "قم", "کردستان",
    "کرمان", "کرمانشاه", "کهگیلویه و بویراحمد", "گلستان", "گیلان", "لرستان", "مازندران",
    "مرکزی", "هرمزگان", "همدان", "یزد",
]
# spellings the booklet also uses for the same province
PROV_ALIASES = {
    "کهگیلویه و بویر احمد": "کهگیلویه و بویراحمد",
    "کهکیلویه و بویراحمد": "کهگیلویه و بویراحمد",
    "کهکیلویه و بویر احمد": "کهگیلویه و بویراحمد",
    "چهار محال و بختیاری": "چهارمحال و بختیاری",
}

UNI_TYPE_RULES = [
    ("پیام نور", "دانشگاه پیام نور"),
    ("غیرانتفاعی", "مؤسسه/دانشگاه غیرانتفاعی"),
    ("غیر انتفاعی", "مؤسسه/دانشگاه غیرانتفاعی"),
    ("آزاد اسلامی", "دانشگاه آزاد اسلامی"),
    ("فرهنگیان", "دانشگاه فرهنگیان"),
    ("ملی مهارت", "دانشگاه ملی مهارت"),
    ("فنی و حرفه", "دانشگاه فنی و حرفه‌ای"),
    ("غیردولتی", "دانشگاه غیردولتی"),
    ("غیر دولتی", "دانشگاه غیردولتی"),
]

SPLIT = re.compile(r"\s+-\s+")


def clean(s):
    s = s.replace(ZWNJ, "")
    s = re.sub(r"\s*[-–—]\s*", " - ", s)
    return " ".join(s.split()).strip(" -")


def find_province(text):
    t = clean(text)
    for wrong, right in PROV_ALIASES.items():
        t = t.replace(wrong, right)
    m = re.search(r"استان\s+(?:محروم\s+)?(" + "|".join(map(re.escape, PROVINCES)) + r")\b", t)
    if m:
        return m.group(1)
    for p in sorted(PROVINCES, key=len, reverse=True):
        if p in t:
            return p
    return ""


def parse_caption(cap, kind):
    """-> (province, institution, native_note)"""
    c = clean(cap)
    if kind == "بومی":
        return "", "", c
    if c.startswith("دانشگاه پیام نور"):
        return find_province(c), c, ""
    if c.startswith("استان "):
        parts = SPLIT.split(c, 1)
        head = parts[0][len("استان "):].strip()
        rest = parts[1].strip() if len(parts) > 1 else ""
        return find_province("استان " + head) or find_province(head), rest or head, ""
    return find_province(c), c, ""


def uni_type(inst, wing):
    for key, label in UNI_TYPE_RULES:
        if key in inst:
            return label
    if "بهداشت" in wing or "علوم پزشکی" in inst or "پزشکی و خدمات بهداشتی" in inst:
        return "دانشگاه/دانشکده علوم پزشکی (دولتی)"
    return "دانشگاه دولتی (وزارت علوم)"


def split_inst_desc(desc):
    """For بومی tables the left column is «دانشگاه محل تحصیل / توضیحات»."""
    d = clean(desc)
    if not d:
        return "", ""
    parts = SPLIT.split(d)
    return parts[0].strip(), " - ".join(p.strip() for p in parts[1:])


def norm_inst(name):
    """Key used to match one university across sections."""
    n = clean(name)
    n = re.split(r"\(\s*محل تحصیل", n)[0]
    n = re.sub(r"\s*\(.*?\)\s*", " ", n)
    n = SPLIT.split(n)[0]
    n = n.replace("دانشگاه ", "").replace("دانشکده ", "")
    return " ".join(n.split())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", required=True)
    ap.add_argument("--sections", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    raw = json.load(open(a.rows, encoding="utf-8"))
    cfg = json.load(open(a.sections, encoding="utf-8"))
    secs = cfg["sections"]

    def assign_sections(rows):
        """Section per row, walking the booklet in order.

        Page ranges may overlap on ONE boundary page, because a booklet can
        change section mid-page — the وزارت بهداشت -> وزارت علوم switch inside
        the big روزانه block does exactly that. The later section then carries
        "from_caption": the caption of its very first table. Rows before that
        caption stay in the earlier section.
        """
        out, i = [], 0
        for r in rows:
            while i + 1 < len(secs):
                nxt = secs[i + 1]
                if r["page"] > secs[i]["to"]:
                    i += 1
                    continue
                fc = nxt.get("from_caption")
                if (fc and r["page"] >= nxt["from"]
                        and clean(r["caption"]).startswith(clean(fc))):
                    i += 1
                    continue
                break
            s = secs[i]
            if not (s["from"] <= r["page"] <= s["to"]):
                raise SystemExit(f"page {r['page']} falls outside every section in "
                                 f"{a.sections} — fix the page ranges and re-run")
            out.append(s)
        return out

    row_sec = assign_sections(raw)

    # pass 1 — collect province per university from the sections whose caption
    # carries the university's own province
    prov_by_inst, staged = {}, []
    for r, s in zip(raw, row_sec):
        sec, sec_dore = s["name"], s.get("dore")
        kind, wing = s.get("kind", "استانی"), s.get("wing", "")
        sec_nahve = s.get("nahve") or ""
        prov, inst, native = parse_caption(r["caption"], kind)
        desc = clean(r["desc"])
        if kind == "بومی":
            inst, desc = split_inst_desc(r["desc"])
        if prov and inst:
            prov_by_inst.setdefault(norm_inst(inst), collections.Counter())[prov] += 1
        staged.append((r, sec, sec_dore, sec_nahve, kind, wing, prov, inst, native, desc))
    prov_lookup = {k: c.most_common(1)[0][0] for k, c in prov_by_inst.items()}
    for name, prov in (cfg.get("province_overrides") or {}).items():
        prov_lookup[norm_inst(name)] = prov

    out, unresolved = [], collections.Counter()
    for r, sec, sec_dore, sec_nahve, kind, wing, prov, inst, native, desc in staged:
        if not prov and inst:
            prov = prov_lookup.get(norm_inst(inst), "")
            if not prov:                     # longest known name contained in it
                flat = clean(inst)
                best = max((k for k in prov_lookup if k and k in flat), key=len, default="")
                prov = prov_lookup.get(best, "")
            if not prov:
                unresolved[inst] += 1

        dore = r["dore"] or sec_dore or ""
        dore_src = "ستون دفترچه" if r["dore"] else "استنتاج از بخش دفترچه"

        zan, mard = r["zan"], r["mard"]
        zan_ok, mard_ok = zan not in ("-", ""), mard not in ("-", "")
        gender = ("زن و مرد" if zan_ok and mard_ok else
                  "فقط زن" if zan_ok else "فقط مرد" if mard_ok else "")

        c1 = int(r["cap1"]) if r["cap1"].isdigit() else 0
        c2 = int(r["cap2"]) if r["cap2"].isdigit() else 0

        out.append({
            "کد رشته محل": r["code"],
            "عنوان رشته": clean(r["title"]),
            "نحوه پذیرش": r["nahve"] or sec_nahve,
            "دوره تحصیلی": dore,
            "جنس پذیرش": gender,
            "استان": prov,
            "دانشگاه / مؤسسه": inst,
            "نوع دانشگاه": uni_type(inst, wing),
            "زیرمجموعه": wing,
            "ظرفیت نیمسال اول": c1,
            "ظرفیت نیمسال دوم": c2,
            "ظرفیت کل": c1 + c2,
            "ظرفیت زن": int(zan) if zan.isdigit() else "",
            "ظرفیت مرد": int(mard) if mard.isdigit() else "",
            "شرط بومی / سهمیه": native,
            "توضیحات": desc,
            "بخش دفترچه": sec,
            "منبع دوره تحصیلی": dore_src,
            "صفحه PDF": r["page"],
        })

    # A booklet may print the same کدرشته‌محل twice (تجربی ۱۴۰۳ repeats 12028-12030
    # in both the وزارت بهداشت tables and the کاردانی عمومی section). The
    # deliverable is one row per code, so drop a repeat only when it says the
    # same thing; a code that reappears with DIFFERENT values is kept, because
    # that would be a real conflict for the reader to see.
    same = lambda r: (r["عنوان رشته"], r["ظرفیت نیمسال اول"], r["ظرفیت نیمسال دوم"],
                      r["جنس پذیرش"], r["دانشگاه / مؤسسه"])
    seen, deduped, repeats = {}, [], 0
    for r in out:
        k = r["کد رشته محل"]
        if k in seen and same(seen[k]) == same(r):
            repeats += 1
            continue
        seen.setdefault(k, r)
        deduped.append(r)
    out = deduped

    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    print("rows:", len(out))
    print("identical repeats dropped:", repeats)
    print("rows with no province:", sum(1 for r in out if not r["استان"]))
    print("unresolved institutions:", len(unresolved))
    for k, v in unresolved.most_common(20):
        print(f"   {v:5d}  {k}")
    print("wrote", a.out)


if __name__ == "__main__":
    main()
