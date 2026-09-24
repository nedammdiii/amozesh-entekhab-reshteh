# -*- coding: utf-8 -*-
"""Re-read the produced workbook and diff every cell against final.json.

openpyxl writing a cell is not proof the cell holds what you meant. Run this
before handing the file over — it caught a real bug on the first booklet.

  python verify_xlsx.py --final final.json --xlsx book.xlsx
"""
import argparse, io, json, sys
from openpyxl import load_workbook

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--final", required=True)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--sheet", default="رشته محل ها")
    a = ap.parse_args()

    rows = json.load(open(a.final, encoding="utf-8"))
    wb = load_workbook(a.xlsx)
    print("sheets:", wb.sheetnames)
    ws = wb[a.sheet]
    print(f"dims: {ws.max_row} x {ws.max_column}")

    hdr = [c.value for c in ws[1]]
    print("headers:", hdr)

    bad = 0
    if ws.max_row != len(rows) + 1:
        print(f"ROW COUNT MISMATCH: sheet {ws.max_row - 1} vs json {len(rows)}")
        bad += 1
    for i, r in enumerate(rows, start=2):
        for j, h in enumerate(hdr, start=1):
            want = r.get(h)
            if want == "":
                want = None
            got = ws.cell(row=i, column=j).value
            if got != want:
                if bad < 5:
                    print(f"  MISMATCH row {i} col {h!r}: {got!r} != {want!r}")
                bad += 1
    print(f"cell mismatches: {bad}")

    formulas = 0
    for name in wb.sheetnames:
        for row in wb[name].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    formulas += 1
    print(f"formula cells: {formulas} (they compute when Excel opens the file)")
    print("PASS" if bad == 0 else "FAIL")
    sys.exit(0 if bad == 0 else 1)


if __name__ == "__main__":
    main()
