"""Sanity-check a built workbook: expected sheets exist and no formula text contains an error literal."""
import sys

import openpyxl

EXPECTED = ["DASHBOARD", "FOOD_DB", "DAILY_INPUT", "SUPPLEMENTS", "ACTIVITY_TDEE", "MICROS",
            "BODY_LOG", "HABITS", "SETTINGS", "SCENARIOS"]
BAD = ("#REF", "#NAME", "#DIV/0", "#VALUE", "#N/A")


def main(path):
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == EXPECTED, wb.sheetnames
    formulas, problems = 0, []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    formulas += c.value.startswith("=")
                    if any(b in c.value for b in BAD):
                        problems.append("%s!%s" % (ws.title, c.coordinate))
    if problems:
        sys.exit("Error literals in: %s" % ", ".join(problems[:20]))
    print("OK: %d sheets, %d formulas, no error literals" % (len(wb.sheetnames), formulas))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "output/MegaTracker.xlsx")
