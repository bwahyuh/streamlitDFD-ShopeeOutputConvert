"""Report Sales dari hasil gabungan Order + Income.

Definisi (per baris produk Order yang sudah ada di Income):
  Value Gross            = Harga Awal x Jumlah                      (harga etalase sebelum diskon)
  Diskon Penjual         = Value Gross - Income Harga Produk          (diskon toko / harga coret)
  Voucher & Diskon Lain  = -(voucher/koin penjual + diskon produk dari Shopee)
  Biaya Marketplace      = -(biaya admin, proses, pembayaran, gratis ongkir XTRA, promo XTRA, AMS, PPh 22, dst.)
  Ongkir Ditanggung Penjual = -(ongkir dibayar pembeli + ongkir ke jasa kirim + subsidi ongkir Shopee + ongkir retur ...)
  Refund                 = -(Jumlah Pengembalian Dana ke Pembeli)
  Penyesuaian Lain       = sisa selisih dari Shopee (mis. penyesuaian kecil pada pesanan refund)
  Total Diskon           = jumlah 6 komponen di atas
  Value Nett             = Value Gross - Total Diskon = Income "Total Penghasilan" (dana yang diterima)

Bundling / dus dipecah ke produk dasar (Tongkol 250gr, Tongkol 500gr, Tuna Can 165gr, ...).
Nilai rupiah satu baris bundling dibagi ke produk dasarnya secara proporsional terhadap harga standar.
"""
from __future__ import annotations

import re

import pandas as pd

KEY = "No. Pesanan"
INC = "Income | "
MARKETPLACE = "Shopee"

T250, T500, TUNA, TUNA1KG, PRJ_MIX = "Tongkol 250gr", "Tongkol 500gr", "Tuna Can 165gr", "Tuna Pouch 1kg", "Paket PRJ Mix 1"
UNMAPPED = "Tidak terpetakan"
PRODUCT_ORDER = [T250, T500, TUNA, TUNA1KG, PRJ_MIX, UNMAPPED]
# Harga standar per pcs: hanya dipakai sebagai bobot pembagian nilai bundling & cek mapping.
STD_PRICE = {T250: 32000, T500: 58000, TUNA: 24500, TUNA1KG: 115000, PRJ_MIX: 129999}
DUS_PCS = {T250: 25, T500: 12, TUNA: 24}

VOUCHER_COLS = ["Voucher disponsor oleh Penjual", "Cashback Koin disponsori Penjual", "Diskon Produk dari Shopee",
                "Voucher co-fund disponsor oleh Penjual", "Cashback Koin Co-fund disponsori Penjual"]
FEE_COLS = ["Biaya Administrasi (termasuk PPN 11%)", "Biaya Proses Pesanan", "Biaya Pembayaran",
            "Biaya Gratis Ongkir XTRA - Ukuran Khusus (Kategori E)", "Biaya Gratis Ongkir XTRA - Ukuran Biasa (Kategori E)",
            "Biaya Transaksi", "Biaya Layanan Promo XTRA", "Biaya Kampanye", "Biaya Komisi AMS",
            "Biaya Isi Saldo Otomatis (dari Penghasilan)", "Premi", "FBS Fee", "PPh 22"]
ONGKIR_COLS = ["Ongkir Dibayar Pembeli", "Ongkos Kirim yang Dibayarkan ke Jasa Kirim", "Potongan Ongkos Kirim dari Jasa Kirim",
               "Gratis Ongkir dari Shopee", "Ongkos Kirim Pengembalian Barang", "Return to Seller Fee", "Pengembalian Biaya Kirim"]
REFUND_COLS = ["Jumlah Pengembalian Dana ke Pembeli"]

DISKON_PARTS = ["Diskon Penjual", "Voucher & Diskon Lain", "Biaya Marketplace",
                "Ongkir Ditanggung Penjual", "Refund", "Penyesuaian Lain"]
MONEY = ["Value Gross"] + DISKON_PARTS + ["Total Diskon", "Value Nett"]


# --------------------------------------------------------------------------- product mapping
def _seg_product(seg: str) -> str | None:
    if "tuna" in seg and "tongkol" not in seg:
        return TUNA1KG if re.search(r"1\s*kg", seg) else TUNA
    if "500" in seg:
        return T500
    if "250" in seg:
        return T250
    if "tuna" in seg or "165" in seg:
        return TUNA
    return None


def parse_product(name: str) -> list[tuple[str, int]]:
    """Nama Produk Shopee -> [(produk dasar, pcs per 1 unit yang dibeli)]."""
    n = re.sub(r"\s+", " ", str(name).lower()).strip()
    n = re.sub(r"^(deltafoods|paket prj)\s*-\s*", "", n)
    if "paket mix 1" in n:
        return [(PRJ_MIX, 1)]
    if "paket spesial delfi" in n:
        return [(T250, 1), (TUNA, 1)]
    if "tongkol dan tuna" in n:                       # "Tongkol dan Tuna Bundling 1" = 250gr + 500gr + Tuna
        return [(T250, 1), (T500, 1), (TUNA, 1)]
    main = re.split(r"\bfree\b", n)[0]                 # hadiah (free ...) tidak dihitung
    segs = [s.strip() for s in re.split(r"\+|\bdan\b", main) if s.strip()]
    if len(segs) > 1:
        out = []
        for s in segs:
            p = _seg_product(s)
            if p is None:
                return []
            m = re.match(r"(\d+)\s", s)
            out.append((p, int(m.group(1)) if m else 1))
        return out
    p = _seg_product(main)
    if p is None:
        return []
    if re.search(r"\bdus\b", main) and p in DUS_PCS:
        return [(p, DUS_PCS[p])]
    m = re.search(r"bundling (\d+)", main) or re.match(r"(\d+)\s", main)
    return [(p, int(m.group(1)) if m else 1)]


def mapping_table(gab: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (name, harga), g in gab.groupby(["Nama Produk", "Harga Awal"], dropna=False):
        comps = parse_product(name)
        std = sum(STD_PRICE[p] * q for p, q in comps)
        ratio = (harga / std) if std and pd.notna(harga) else None
        status = "Tidak terpetakan" if not comps else ("OK" if ratio and 0.9 <= ratio <= 1.15 else "Cek: harga beda jauh dari harga standar")
        rows.append({
            "Nama Produk (Shopee)": name,
            "Harga Awal": harga,
            "Komposisi": " + ".join(f"{q} x {p}" for p, q in comps) or "-",
            "Total pcs": sum(q for _, q in comps),
            "Nilai standar": std or None,
            "Harga Awal / standar": ratio,
            "Baris": len(g),
            "Status": status,
        })
    return pd.DataFrame(rows).sort_values(["Status", "Nama Produk (Shopee)"], ascending=[False, True]).reset_index(drop=True)


# --------------------------------------------------------------------------- detail
def _neg_sum(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    present = [INC + c for c in cols if INC + c in df.columns]
    if not present:
        return pd.Series(0.0, index=df.index)
    return -df[present].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)


def build_detail(gab: pd.DataFrame) -> pd.DataFrame:
    """1 baris = 1 produk dasar dalam 1 baris Order yang sudah ada di Income."""
    g = gab[gab["Status Match Income"].astype(str).str.startswith("Cocok")].copy()
    if g.empty:
        return pd.DataFrame(columns=["MARKETPLACE", "Bulan", "Tanggal Pesanan", KEY, "PRODUK", "Qty Produk", "Qty Retur"] + MONEY)

    num = lambda c: pd.to_numeric(g[c], errors="coerce").fillna(0)  # noqa: E731
    g["_gross"] = num("Harga Awal") * num("Jumlah")
    g["_harga_produk"] = pd.to_numeric(g[INC + "Harga Produk"], errors="coerce").fillna(0)
    g["_voucher"] = _neg_sum(g, VOUCHER_COLS)
    g["_biaya"] = _neg_sum(g, FEE_COLS)
    g["_ongkir"] = _neg_sum(g, ONGKIR_COLS)
    g["_refund"] = _neg_sum(g, REFUND_COLS)
    g["_nett"] = pd.to_numeric(g[INC + "Total Penghasilan"], errors="coerce").fillna(0)
    parts = ["_harga_produk", "_voucher", "_biaya", "_ongkir", "_refund", "_nett"]

    # Pesanan multi-produk tanpa rincian Sku: nilai pesanan (ada di baris pertama) dibagi ke baris
    # produknya secara proporsional terhadap Value Gross.
    multi = g["Status Match Income"].str.startswith("Cocok per pesanan (multi")
    if multi.any():
        gm = g[multi]
        share = gm["_gross"] / gm.groupby(KEY)["_gross"].transform("sum").replace(0, pd.NA)
        totals = gm.groupby(KEY)[parts].transform("sum")
        g.loc[multi, parts] = totals.mul(share.fillna(0), axis=0).to_numpy()

    g["_diskon_penjual"] = g["_gross"] - g["_harga_produk"]
    g["_lain"] = (g["_gross"] - g["_diskon_penjual"] - g["_voucher"] - g["_biaya"] - g["_ongkir"]
                  - g["_refund"] - g["_nett"])
    qty = num("Jumlah")
    ret = num("Returned quantity") if "Returned quantity" in g else 0 * qty
    tgl = pd.to_datetime(g["Waktu Pesanan Dibuat"], errors="coerce")

    rows = []
    for i in range(len(g)):
        comps = parse_product(g["Nama Produk"].iat[i]) or [(UNMAPPED, 1)]
        std_total = sum(STD_PRICE.get(p, 1) * q for p, q in comps)
        for p, q in comps:
            w = STD_PRICE.get(p, 1) * q / std_total
            rows.append({
                "MARKETPLACE": MARKETPLACE,
                "Bulan": tgl.iat[i].strftime("%Y-%m") if pd.notna(tgl.iat[i]) else "",
                "Tanggal Pesanan": tgl.iat[i],
                KEY: g[KEY].iat[i],
                "Nama Produk (Shopee)": g["Nama Produk"].iat[i],
                "PRODUK": p,
                "Qty Produk": q * (qty.iat[i] - ret.iat[i]),
                "Qty Retur": q * ret.iat[i],
                "Value Gross": g["_gross"].iat[i] * w,
                "Diskon Penjual": g["_diskon_penjual"].iat[i] * w,
                "Voucher & Diskon Lain": g["_voucher"].iat[i] * w,
                "Biaya Marketplace": g["_biaya"].iat[i] * w,
                "Ongkir Ditanggung Penjual": g["_ongkir"].iat[i] * w,
                "Refund": g["_refund"].iat[i] * w,
                "Penyesuaian Lain": g["_lain"].iat[i] * w,
                "Value Nett": g["_nett"].iat[i] * w,
            })
    d = pd.DataFrame(rows)
    d["Total Diskon"] = d[DISKON_PARTS].sum(axis=1)
    cols = ["MARKETPLACE", "Bulan", "Tanggal Pesanan", KEY, "Nama Produk (Shopee)", "PRODUK", "Qty Produk", "Qty Retur"] + MONEY
    return d[cols]


# --------------------------------------------------------------------------- pivot
def pivot(detail: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame(columns=by + ["Jumlah Pesanan", "Qty Produk", "Qty Retur"] + MONEY + ["% Total Diskon", "Nett per Pcs"])
    agg = detail.groupby(by, sort=False).agg(
        **{"Jumlah Pesanan": (KEY, "nunique"), "Qty Produk": ("Qty Produk", "sum"), "Qty Retur": ("Qty Retur", "sum")},
        **{c: (c, "sum") for c in MONEY},
    ).reset_index()
    if "PRODUK" in by:
        agg["_o"] = agg["PRODUK"].map({p: i for i, p in enumerate(PRODUCT_ORDER)}).fillna(99)
        agg = agg.sort_values([c for c in by if c != "PRODUK"] + ["_o"]).drop(columns="_o")
    total = {c: "" for c in by}
    total[by[0]] = "TOTAL"
    total.update({"Jumlah Pesanan": detail[KEY].nunique(), "Qty Produk": detail["Qty Produk"].sum(),
                  "Qty Retur": detail["Qty Retur"].sum(), **{c: detail[c].sum() for c in MONEY}})
    agg = pd.concat([agg, pd.DataFrame([total])], ignore_index=True)
    agg["% Total Diskon"] = (agg["Total Diskon"] / agg["Value Gross"].replace(0, pd.NA)).astype(float)
    agg["Nett per Pcs"] = (agg["Value Nett"] / agg["Qty Produk"].replace(0, pd.NA)).astype(float)
    return agg


def build_report(gab: pd.DataFrame) -> dict[str, pd.DataFrame]:
    detail = build_detail(gab)
    return {
        "detail": detail,
        "by_produk": pivot(detail, ["MARKETPLACE", "PRODUK"]),
        "by_bulan": pivot(detail, ["MARKETPLACE", "Bulan", "PRODUK"]),
        "mapping": mapping_table(gab),
    }


# --------------------------------------------------------------------------- excel
def write_report_sheets(xw, gab: pd.DataFrame) -> None:
    """Tulis sheet 'Report Sales', 'Data Sales', 'Mapping Produk' ke ExcelWriter (openpyxl)."""
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    rep = build_report(gab)
    bold = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="DDEBF7")
    total_fill = PatternFill("solid", fgColor="F2F2F2")

    def fmt_table(ws, df: pd.DataFrame, startrow: int) -> None:
        hr = startrow + 1  # header row (1-based)
        for j, c in enumerate(df.columns, 1):
            cell = ws.cell(row=hr, column=j)
            cell.font, cell.fill = bold, head_fill
            fmt = "0.0%" if c == "% Total Diskon" else ("#,##0" if c in MONEY + ["Qty Produk", "Qty Retur", "Jumlah Pesanan", "Nett per Pcs"] else None)
            for i in range(len(df)):
                cell = ws.cell(row=hr + 1 + i, column=j)
                if fmt:
                    cell.number_format = fmt
                if df.iloc[i, 0] == "TOTAL":
                    cell.font, cell.fill = bold, total_fill

    # Report Sales
    sheet = "Report Sales"
    pd.DataFrame().to_excel(xw, sheet_name=sheet)
    ws = xw.sheets[sheet]
    ws["A1"] = "REPORT SALES SHOPEE"
    ws["A1"].font = Font(bold=True, size=14)
    det = rep["detail"]
    periode = f"{det['Tanggal Pesanan'].min():%d %b %Y} - {det['Tanggal Pesanan'].max():%d %b %Y}" if len(det) else "-"
    ws["A2"] = f"Periode pesanan dibuat: {periode}. Hanya pesanan yang dananya sudah dilepas (ada di Income)."
    ws["A3"] = ("Value Gross = Harga Awal x Jumlah. Value Nett = dana diterima (Income 'Total Penghasilan'). "
                "Total Diskon = Value Gross - Value Nett = " + " + ".join(DISKON_PARTS) + ".")
    row = 5
    ws.cell(row=row, column=1, value="Per Produk").font = Font(bold=True, size=12)
    rep["by_produk"].to_excel(xw, sheet_name=sheet, startrow=row, index=False)
    fmt_table(ws, rep["by_produk"], row)
    row += len(rep["by_produk"]) + 4
    ws.cell(row=row, column=1, value="Per Bulan & Produk").font = Font(bold=True, size=12)
    rep["by_bulan"].to_excel(xw, sheet_name=sheet, startrow=row, index=False)
    fmt_table(ws, rep["by_bulan"], row)
    widths = [14, 16, 16] + [16] * 20
    for j, w in enumerate(widths[: max(len(rep["by_bulan"].columns), 1)], 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # Data Sales (detail, bisa dibuat Pivot Table sendiri di Excel)
    det.to_excel(xw, sheet_name="Data Sales", index=False)
    ws = xw.sheets["Data Sales"]
    ws.freeze_panes = "A2"
    if len(det.columns):
        ws.auto_filter.ref = ws.dimensions
    fmt_table(ws, det, 0)
    for j, c in enumerate(det.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = 45 if c == "Nama Produk (Shopee)" else 16
        if c == "Tanggal Pesanan":
            for i in range(len(det)):
                ws.cell(row=2 + i, column=j).number_format = "yyyy-mm-dd hh:mm"

    # Mapping Produk
    mp = rep["mapping"]
    mp.to_excel(xw, sheet_name="Mapping Produk", index=False)
    ws = xw.sheets["Mapping Produk"]
    ws.freeze_panes = "A2"
    for j, c in enumerate(mp.columns, 1):
        ws.cell(row=1, column=j).font = bold
        ws.column_dimensions[get_column_letter(j)].width = 90 if c == "Nama Produk (Shopee)" else (40 if c in ("Komposisi", "Status") else 14)
        nf = "0.00" if c == "Harga Awal / standar" else ("#,##0" if c in ("Harga Awal", "Nilai standar") else None)
        if nf:
            for i in range(len(mp)):
                ws.cell(row=2 + i, column=j).number_format = nf
