"""Gabungkan export Shopee "Order.all" dengan "Income" (Laporan Penghasilan).

Semua fungsi menerima path atau file-like object (BytesIO dari st.file_uploader),
sehingga modul ini bisa dipakai ulang di aplikasi Streamlit tanpa menulis ke disk.

Aturan penggabungan
-------------------
* Kunci: No. Pesanan.
* 1 baris hasil = 1 baris produk di file Order (struktur file Order dipertahankan).
* Kolom Income diambil dari baris "Sku" (rincian per produk) yang cocok dengan
  (No. Pesanan, Nama Produk). Jika pesanan tidak punya baris Sku dan hanya punya
  1 produk, dipakai baris "Order" (nilainya sama karena hanya 1 produk).
* Kolom Income yang hanya ada di level pesanan (kosong / "-" di baris Sku):
  - kolom teks  -> diisi di semua baris produk pesanan tsb.
  - kolom angka -> diisi hanya di baris produk pertama pesanan (agar SUM tidak dobel).
* Sheet "Seller Fee" & "Shipping Fee Discrepancy" juga level pesanan -> baris pertama saja.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

KEY = "No. Pesanan"
INC = "Income | "
FEE = "Seller Fee | "
SHIP = "Selisih Ongkir | "

# Kolom angka di file Order memakai pemisah ribuan titik ("58.000").
_ORDER_NUM_RE = re.compile(r"^-?\d{1,3}(\.\d{3})*$")
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}(:\d{2})?)?$")
_BLANK = {"", "-", "nan", "None"}


def _norm_name(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().str.replace(r"\s+", " ", regex=True).str.lower()


def _is_blank(s: pd.Series) -> pd.Series:
    return s.isna() | s.astype(str).str.strip().isin(_BLANK)


def _convert_order_types(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in df.columns:
        if c == KEY:
            continue
        vals = df[c].dropna().astype(str).str.strip()
        vals = vals[vals != ""]
        if vals.empty:
            continue
        if vals.str.match(_ORDER_NUM_RE).all():
            df[c] = pd.to_numeric(df[c].astype(str).str.replace(".", "", regex=False), errors="coerce")
        elif vals.str.match(_DATE_RE).all():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def _convert_income_types(df: pd.DataFrame, skip: set[str]) -> pd.DataFrame:
    """Angka di file Income tanpa pemisah ribuan; '-' / '' = kosong."""
    df = df.copy()
    for c in df.columns:
        if c in skip:
            continue
        vals = df[c].dropna().astype(str).str.strip()
        vals = vals[~vals.isin(_BLANK)]
        if vals.empty:
            continue
        if pd.to_numeric(vals, errors="coerce").notna().all():
            df[c] = pd.to_numeric(df[c].astype(str).str.strip().replace(list(_BLANK), pd.NA), errors="coerce")
        elif vals.str.match(_DATE_RE).all():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


# --------------------------------------------------------------------------- helpers
BULAN = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
_MONTH_WORDS = {
    1: ["januari", "january", "jan"],
    2: ["februari", "pebruari", "february", "feb", "peb"],
    3: ["maret", "march", "mar"],
    4: ["april", "apr"],
    5: ["mei", "may"],
    6: ["juni", "june", "jun"],
    7: ["juli", "july", "jul"],
    8: ["agustus", "august", "agu", "agt", "ags", "aug"],
    9: ["september", "sept", "sep"],
    10: ["oktober", "october", "okt", "oct"],
    11: ["november", "nopember", "nov", "nop"],
    12: ["desember", "december", "des", "dec"],
}
_MONTHS = {w: m for m, words in _MONTH_WORDS.items() for w in words}


@dataclass(frozen=True)
class Period:
    year: int | None
    month: int

    @property
    def label(self) -> str:
        return f"{BULAN[self.month]} {self.year}" if self.year else BULAN[self.month]

    def matches(self, other: "Period") -> bool:
        """Bulan sama; tahun dibandingkan hanya jika keduanya tercantum di nama file."""
        return self.month == other.month and (self.year is None or other.year is None or self.year == other.year)


def parse_period(filename: str) -> Period | None:
    """Ambil bulan (dan tahun bila ada) dari NAMA file.

    'Order.all Shopee JULI.xlsx'            -> Juli
    'Income.sudah dilepas.id AGUSTUS.xlsx'  -> Agustus
    'Order.all.20260601_20260630.xlsx'      -> Juni 2026 (tanggal pertama)
    'Order Okt 2026.xlsx'                   -> Oktober 2026
    None jika tidak ada bulan, atau ada lebih dari satu bulan berbeda di nama file.
    """
    stem = re.sub(r"\.xls[xm]?$", "", filename.strip(), flags=re.I)
    m = re.search(r"(20\d{2})(0[1-9]|1[0-2])([0-3]\d)", stem)
    if m:
        return Period(int(m.group(1)), int(m.group(2)))
    tokens = re.findall(r"[a-z]+|\d+", stem.lower())
    months = {_MONTHS[t] for t in tokens if t in _MONTHS}
    if len(months) != 1:
        return None
    year = next((int(t) for t in tokens if re.fullmatch(r"20\d{2}", t)), None)
    return Period(year, months.pop())


def detect_file_type(src) -> str | None:
    """'order' jika ada sheet 'orders', 'income' jika ada sheet 'Penghasilan', selain itu None."""
    try:
        sheets = pd.ExcelFile(src).sheet_names
    except Exception:
        return None
    if "orders" in sheets:
        return "order"
    if "Penghasilan" in sheets:
        return "income"
    return None


# --------------------------------------------------------------------------- readers
def read_order(src, name: str = "") -> pd.DataFrame:
    df = pd.read_excel(src, sheet_name="orders", dtype=str)
    df = df[df[KEY].notna()].copy()
    df[KEY] = df[KEY].str.strip()
    df = _convert_order_types(df)
    df.insert(0, "Sumber File Order", name)
    return df


@dataclass
class Income:
    penghasilan: pd.DataFrame          # semua baris (Order + Sku), kolom asli
    seller_fee: pd.DataFrame
    shipping_discrepancy: pd.DataFrame
    name: str = ""


def read_income(src, name: str = "") -> Income:
    xl = pd.ExcelFile(src)
    p = pd.read_excel(xl, sheet_name="Penghasilan", header=2, dtype=str)
    p = p[p[KEY].notna()].copy()
    p[KEY] = p[KEY].str.strip()
    text_cols = {KEY, "No.", "Lihat berdasarkan", "No. Pengajuan", "ID Produk", "Nama Produk",
                 "Username (Pembeli)", "Kode Voucher", "Rencana Cicilan (jika berlaku)"}
    p = _convert_income_types(p, skip=text_cols)
    p["Sumber File Income"] = name

    sf = pd.DataFrame(columns=[KEY])
    if "Seller Fee" in xl.sheet_names:
        sf = pd.read_excel(xl, sheet_name="Seller Fee", header=1, dtype=str)
        sf = sf[sf[KEY].notna()].drop(columns=["No."], errors="ignore")
        sf[KEY] = sf[KEY].str.strip()
        sf = _convert_income_types(sf, skip={KEY})

    sd = pd.DataFrame(columns=[KEY])
    if "Shipping Fee Discrepancy" in xl.sheet_names:
        sd = pd.read_excel(xl, sheet_name="Shipping Fee Discrepancy", header=1, dtype=str)
        sd = sd[sd[KEY].notna()].copy()
        sd[KEY] = sd[KEY].str.strip()
        sd = _convert_income_types(sd, skip={KEY, "Alasan perbedaan"})
    return Income(p, sf, sd, name)


# --------------------------------------------------------------------------- merge
@dataclass
class MergeResult:
    gabungan: pd.DataFrame
    income_tanpa_order: pd.DataFrame
    ringkasan: pd.DataFrame
    catatan: list[str] = field(default_factory=list)


def _concat_income(incomes: list[Income]):
    p = pd.concat([i.penghasilan for i in incomes], ignore_index=True)
    sf = pd.concat([i.seller_fee for i in incomes], ignore_index=True).drop_duplicates(KEY, keep="last")
    sd = pd.concat([i.shipping_discrepancy for i in incomes], ignore_index=True).drop_duplicates(KEY, keep="last")
    return p, sf, sd


def merge(orders: list[pd.DataFrame], incomes: list[Income]) -> MergeResult:
    """orders: hasil read_order (boleh lebih dari 1 periode). incomes: hasil read_income (semua periode).

    Income digabung dari semua periode, karena dana pesanan bulan X bisa dilepas di bulan X+1.
    """
    order = pd.concat(orders, ignore_index=True)
    order["_line"] = range(len(order))
    order["_name"] = _norm_name(order["Nama Produk"])
    order["_first"] = ~order.duplicated(KEY)

    p, sf, sd = _concat_income(incomes)
    inc_cols = [c for c in p.columns if c not in {"No.", "Lihat berdasarkan", KEY}]
    ord_rows = p[p["Lihat berdasarkan"] == "Order"].drop_duplicates(KEY, keep="last").set_index(KEY)
    sku_rows = p[p["Lihat berdasarkan"] == "Sku"].copy()
    sku_rows["_name"] = _norm_name(sku_rows["Nama Produk"])

    # Kolom level-pesanan = kolom yang di baris Sku selalu kosong / "-".
    order_level = [c for c in inc_cols if c != "Sumber File Income" and _is_blank(sku_rows[c]).all()] if len(sku_rows) else []
    numeric_order_level = [c for c in order_level if pd.api.types.is_numeric_dtype(ord_rows[c])]
    text_order_level = [c for c in order_level if c not in numeric_order_level]

    # 1) cocokkan baris produk ke baris Sku: (No. Pesanan, Nama Produk) + urutan kemunculan
    order["_k"] = order.groupby([KEY, "_name"]).cumcount()
    sku_rows["_k"] = sku_rows.groupby([KEY, "_name"]).cumcount()
    m = order[[KEY, "_name", "_k", "_line"]].merge(
        sku_rows[[KEY, "_name", "_k"] + inc_cols], on=[KEY, "_name", "_k"], how="left")
    m = m.set_index("_line").reindex(order["_line"])
    match = pd.Series("", index=order.index)
    has_sku = m["Sumber File Income"].notna().to_numpy()
    match[has_sku] = "Cocok per produk (Sku)"

    inc_block = m[inc_cols].reset_index(drop=True)

    # 2) fallback: pesanan 1 produk tanpa baris Sku -> pakai baris Order
    n_lines = order.groupby(KEY)[KEY].transform("size")
    in_ord = order[KEY].isin(ord_rows.index)
    fb = (~has_sku) & in_ord.to_numpy() & (n_lines == 1).to_numpy()
    if fb.any():
        src = ord_rows.loc[order.loc[fb, KEY], inc_cols].to_numpy()
        inc_block.loc[fb, inc_cols] = src
        match[fb] = "Cocok per pesanan (1 produk)"
    # 2b) pesanan >1 produk tanpa baris Sku sama sekali -> nilai baris Order ditaruh di baris pertama,
    #     kolom teks diisi di semua baris (total tetap benar, rincian per produk tidak tersedia).
    sku_ids = set(sku_rows[KEY])
    no_sku_multi = (~has_sku) & in_ord.to_numpy() & (n_lines > 1).to_numpy() & ~order[KEY].isin(sku_ids).to_numpy()
    if no_sku_multi.any():
        text_cols = [c for c in inc_cols if not pd.api.types.is_numeric_dtype(p[c])]
        first = no_sku_multi & order["_first"].to_numpy()
        inc_block.loc[first, inc_cols] = ord_rows.loc[order.loc[first, KEY], inc_cols].to_numpy()
        others = no_sku_multi & ~order["_first"].to_numpy()
        inc_block.loc[others, text_cols] = ord_rows.loc[order.loc[others, KEY], text_cols].to_numpy()
        match[no_sku_multi] = "Cocok per pesanan (multi produk, angka di baris pertama)"
    rest = (~has_sku) & (~fb) & (~no_sku_multi) & in_ord.to_numpy()
    match[rest] = "Ada di Income, produk tidak cocok"
    match[(match == "")] = "Belum ada di Income"

    # 3) kolom level-pesanan
    ok = in_ord.to_numpy()
    for c in text_order_level:
        inc_block.loc[ok, c] = ord_rows.loc[order.loc[ok, KEY], c].to_numpy()
    first_ok = ok & order["_first"].to_numpy()
    for c in numeric_order_level:
        inc_block[c] = pd.NA
        inc_block.loc[first_ok, c] = ord_rows.loc[order.loc[first_ok, KEY], c].to_numpy()
        inc_block[c] = pd.to_numeric(inc_block[c], errors="coerce")
    for c in inc_cols:
        if c in numeric_order_level:
            continue
        if pd.api.types.is_numeric_dtype(p[c]):
            inc_block[c] = pd.to_numeric(inc_block[c], errors="coerce")
        elif pd.api.types.is_datetime64_any_dtype(p[c]):
            inc_block[c] = pd.to_datetime(inc_block[c], errors="coerce")

    inc_block.columns = [INC + c if c != "Sumber File Income" else c for c in inc_cols]

    # 4) Seller Fee & Selisih Ongkir (level pesanan, baris pertama saja)
    def order_level_block(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
        cols = [c for c in df.columns if c != KEY]
        out = pd.DataFrame(index=order.index, columns=[prefix + c for c in cols])
        d = df.drop_duplicates(KEY).set_index(KEY)
        sel = order["_first"] & order[KEY].isin(d.index)
        for c in cols:
            vals = d.loc[order.loc[sel, KEY], c].to_numpy()
            out.loc[sel, prefix + c] = vals
            if pd.api.types.is_numeric_dtype(d[c]):
                out[prefix + c] = pd.to_numeric(out[prefix + c], errors="coerce")
        return out

    fee_block = order_level_block(sf, FEE)
    ship_block = order_level_block(sd, SHIP)

    base = order.drop(columns=["_line", "_name", "_first", "_k"]).reset_index(drop=True)
    gab = pd.concat([base, pd.DataFrame({"Status Match Income": match.to_numpy()}),
                     inc_block, fee_block.reset_index(drop=True), ship_block.reset_index(drop=True)], axis=1)
    src_col = gab.pop("Sumber File Income")
    gab.insert(1, "Sumber File Income", src_col)

    # income yang pesanannya tidak ada di file Order yang diupload
    orphan_ids = set(ord_rows.index) - set(order[KEY])
    orphan = p[p[KEY].isin(orphan_ids)].copy()

    # ringkasan per file Order
    rows = []
    g = gab.drop_duplicates(KEY)
    for f, gg in g.groupby("Sumber File Order", sort=False):
        lines = gab[gab["Sumber File Order"] == f]
        rows.append({
            "File Order": f,
            "Baris produk": len(lines),
            "Jumlah pesanan": len(gg),
            "Pesanan Selesai": (gg["Status Pesanan"] == "Selesai").sum(),
            "Pesanan Batal": (gg["Status Pesanan"] == "Batal").sum(),
            "Pesanan status lain": (~gg["Status Pesanan"].isin(["Selesai", "Batal"])).sum(),
            "Pesanan ada di Income": (gg["Status Match Income"] != "Belum ada di Income").sum(),
            "Pesanan belum ada di Income": (gg["Status Match Income"] == "Belum ada di Income").sum(),
            "Baris cocok per produk (Sku)": (lines["Status Match Income"] == "Cocok per produk (Sku)").sum(),
            "Baris cocok per pesanan (1 produk)": (lines["Status Match Income"] == "Cocok per pesanan (1 produk)").sum(),
            "Baris cocok per pesanan (multi produk)": lines["Status Match Income"].str.startswith("Cocok per pesanan (multi").sum(),
            "Baris produk tidak cocok": (lines["Status Match Income"] == "Ada di Income, produk tidak cocok").sum(),
        })
    ringkasan = pd.DataFrame(rows)

    catatan = [
        "1 baris = 1 baris produk dari file Order (urutan & kolom asli file Order dipertahankan).",
        "Kunci gabung: No. Pesanan. Income dicari di SEMUA file Income yang diupload (dana bisa dilepas bulan berikutnya).",
        "Kolom 'Income | ...' diambil dari baris Sku (rincian per produk) yang cocok dengan Nama Produk; "
        "jika pesanan hanya 1 produk & tidak punya baris Sku, dipakai baris Order.",
        "Jika pesanan >1 produk tapi Income tidak punya rincian Sku, angka dari baris Order ditaruh di baris "
        "produk pertama saja (status 'Cocok per pesanan (multi produk, ...)').",
        "Kolom Income level-pesanan berupa angka hanya diisi di baris produk pertama pesanan (agar SUM tidak dobel): "
        + ", ".join(numeric_order_level),
        "Kolom Income level-pesanan berupa teks diisi di semua baris: " + ", ".join(text_order_level),
        "Kolom 'Seller Fee | ...' dan 'Selisih Ongkir | ...' adalah level pesanan, hanya di baris produk pertama.",
        "Biaya di file Income bernilai negatif (potongan), sesuai file asli Shopee.",
        "Sheet 'Income tanpa Order' = pesanan yang ada di Income tapi file Order-nya tidak diupload (mis. pesanan bulan sebelumnya).",
        "Sheet 'Report Sales' & 'Data Sales' hanya memakai pesanan yang sudah ada di Income (dana sudah dilepas). "
        "Bundling/dus dipecah ke produk dasar sesuai sheet 'Mapping Produk'; nilai rupiah bundling dibagi "
        "proporsional terhadap harga standar produk dasar.",
        "Report: Value Gross = Harga Awal x Jumlah; Value Nett = Income 'Total Penghasilan'; "
        "Total Diskon = Value Gross - Value Nett (rinciannya ada di kolom komponen diskon).",
        "Report: Qty Produk = (Jumlah - Returned quantity) x isi bundling; Qty Retur ditampilkan terpisah.",
    ]
    return MergeResult(gab, orphan, ringkasan, catatan)


# --------------------------------------------------------------------------- writer
def to_excel(result: MergeResult, dest) -> None:
    """Tulis hasil ke path / BytesIO dengan format sederhana (header tebal, filter, freeze)."""
    from openpyxl.utils import get_column_letter

    import sales_report

    with pd.ExcelWriter(dest, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm", date_format="yyyy-mm-dd") as xw:
        # Sheet pertama: Report Sales, lalu Data Sales & Mapping Produk.
        sales_report.write_report_sheets(xw, result.gabungan)
        sheets = {
            "Gabungan": result.gabungan,
            "Ringkasan": result.ringkasan,
            "Income tanpa Order": result.income_tanpa_order,
            "Keterangan": pd.DataFrame({"Keterangan": result.catatan}),
        }
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "B2" if name == "Gabungan" else "A2"
            if len(df.columns):
                ws.auto_filter.ref = ws.dimensions
            for i, c in enumerate(df.columns, 1):
                col = get_column_letter(i)
                ws[f"{col}1"].font = ws[f"{col}1"].font.copy(bold=True)
                width = min(max(len(str(c)), 10) + 2, 45 if name != "Keterangan" else 160)
                ws.column_dimensions[col].width = width
                s = df[c]
                if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s) \
                        and (s.dropna() % 1 == 0).all():
                    for cell in ws[col][1:]:
                        cell.number_format = "#,##0"


def split_by_order_file(result: MergeResult, order_file: str, income_file: str | None) -> MergeResult:
    """Ambil bagian hasil untuk 1 file Order (dan income-tanpa-order dari file Income pasangannya)."""
    gab = result.gabungan[result.gabungan["Sumber File Order"] == order_file].reset_index(drop=True)
    orphan = result.income_tanpa_order
    orphan = orphan[orphan["Sumber File Income"] == income_file] if income_file else orphan.iloc[0:0]
    ring = result.ringkasan[result.ringkasan["File Order"] == order_file]
    return MergeResult(gab, orphan.reset_index(drop=True), ring.reset_index(drop=True), result.catatan)
