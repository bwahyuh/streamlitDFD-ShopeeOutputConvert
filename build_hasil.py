"""Gabungkan semua pasangan file Order + Income di folder "ACUAN INPUT" ke folder "hasil".

Pemakaian:  python build_hasil.py
Output:
  hasil/Gabungan Shopee <Periode>.xlsx      -> 1 file per file Order
  hasil/Gabungan Shopee Semua Periode.xlsx  -> semua periode jadi satu
"""
from pathlib import Path

import shopee_merge as sm

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "ACUAN INPUT"
OUTPUT = ROOT / "hasil"


def main() -> None:
    OUTPUT.mkdir(exist_ok=True)
    order_files = sorted(INPUT.glob("Order*.xlsx"))
    income_files = sorted(INPUT.glob("Income*.xlsx"))

    orders = [sm.read_order(f, f.name) for f in order_files]
    incomes = [sm.read_income(f, f.name) for f in income_files]
    income_periods = [(sm.parse_period(f.name), f.name) for f in income_files]

    result = sm.merge(orders, incomes)

    for f in order_files:
        p = sm.parse_period(f.name)
        label = p.label if p else f.stem
        inc = next((n for ip, n in income_periods if p and ip and p.matches(ip)), None)
        if p and inc is None:
            print(f"PERINGATAN: {f.name} tidak punya file Income bulan {label}")
        part = sm.split_by_order_file(result, f.name, inc)
        out = OUTPUT / f"Gabungan Shopee {label}.xlsx"
        sm.to_excel(part, out)
        print(f"{out.name}: {len(part.gabungan)} baris, {part.gabungan.shape[1]} kolom, "
              f"income tanpa order: {part.income_tanpa_order[sm.KEY].nunique() if len(part.income_tanpa_order) else 0} pesanan")

    out = OUTPUT / "Gabungan Shopee Semua Periode.xlsx"
    sm.to_excel(result, out)
    print(f"{out.name}: {len(result.gabungan)} baris")
    print()
    print(result.ringkasan.to_string(index=False))


if __name__ == "__main__":
    main()
