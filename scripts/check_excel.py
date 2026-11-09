"""Recalculate every formula in the workbook (no Excel/LibreOffice needed) and report errors.

Usage: python check_excel.py workbook.xlsx [--show Sheet!A1 ...]
Exit code 1 if any formula evaluates to an Excel error.
"""
import sys, warnings

import formulas

warnings.filterwarnings("ignore")
ERRORS = ("#DIV/0!", "#REF!", "#NAME?", "#VALUE!", "#N/A", "#NUM!", "#NULL!")


def main():
    path = sys.argv[1]
    show = sys.argv[sys.argv.index("--show") + 1:] if "--show" in sys.argv else []
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()
    n_formula, bad = 0, []
    for k, v in sol.items():
        if not hasattr(v, "value"):
            continue
        val = v.value
        try:
            vals = list(val.ravel())
        except AttributeError:
            vals = [val]
        for x in vals:
            if str(x).upper() in ERRORS or type(x).__name__ == "XlError":
                bad.append((k, str(x)))
                break
    n_formula = sum(1 for c in xl.cells.values() if getattr(c, "inputs", None))
    print(f"公式重算完成：{n_formula} 个公式，错误单元格 {len(bad)} 个")
    for k, x in bad[:40]:
        print("  ", k, x)
    for ref in show:
        sheet, cell = ref.split("!")
        key = next((k for k in sol if k.upper().endswith(f"[{path.split('/')[-1].upper()}]{sheet.upper()}'!{cell.upper()}")), None)
        print(ref, "=", sol[key].value.ravel()[0] if key else "未找到")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
