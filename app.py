"""Streamlit app: gabungkan export Shopee Order.all + Income (Laporan Penghasilan).

Privasi: file hanya dibaca di memori selama sesi. Tidak ada yang ditulis ke disk,
tidak memakai st.cache_data untuk data upload, dan isi data tidak dicetak ke log.
"""
from __future__ import annotations

import io
from dataclasses import dataclass

import pandas as pd
import streamlit as st

import sales_report
import shopee_merge as sm

st.set_page_config(page_title="Shopee Order + Income", page_icon="🧾", layout="wide")


def _arrow_available() -> bool:
    # st.dataframe butuh pyarrow. Di sebagian PC Windows DLL-nya bisa diblokir kebijakan keamanan.
    try:
        import pyarrow  # noqa: F401
        return True
    except Exception:
        return False


ARROW_OK = _arrow_available()


def show_df(df: pd.DataFrame) -> None:
    if ARROW_OK:
        st.dataframe(df, hide_index=True, width="stretch")
    else:
        html = df.to_html(index=False, na_rep="", border=0)
        st.markdown(f'<div style="overflow:auto;max-height:420px;font-size:0.8rem">{html}</div>',
                    unsafe_allow_html=True)


@dataclass
class Upload:
    name: str
    data: bytes
    expected: str              # "order" / "income" = kolom upload yang dipakai
    kind: str | None           # jenis menurut isi file (nama sheet)
    period: sm.Period | None   # bulan menurut NAMA file


@dataclass
class Output:
    label: str
    filename: str
    data: bytes


def fmt_report(df: pd.DataFrame) -> pd.DataFrame:
    """Format angka report untuk tampilan (ribuan pakai titik, persen 1 desimal)."""
    out = df.copy()
    for c in out.columns:
        if c == "% Total Diskon":
            out[c] = out[c].map(lambda v: "" if pd.isna(v) else f"{v * 100:.1f}%".replace(".", ","))
        elif pd.api.types.is_numeric_dtype(out[c]):
            out[c] = out[c].map(lambda v: "" if pd.isna(v) else f"{v:,.0f}".replace(",", "."))
    return out


def pair_income(order: Upload, incomes: list[Upload]) -> Upload | None:
    return next((i for i in incomes if i.kind == "income" and i.period and order.period
                 and order.period.matches(i.period)), None)


def validate(orders: list[Upload], incomes: list[Upload]) -> tuple[list[str], list[str], pd.DataFrame]:
    """Cek jenis file, bulan dari nama file, duplikat, dan pasangan Order <-> Income per bulan."""
    errors: list[str] = []
    notes: list[str] = []

    for u in orders:
        if u.kind == "income":
            errors.append(f"`{u.name}` diupload di kolom **Order**, tapi isinya file **Income**. Pindahkan ke kolom Income.")
        elif u.kind != "order":
            errors.append(f"`{u.name}` bukan export Order Shopee (tidak ada sheet `orders`).")
    for u in incomes:
        if u.kind == "order":
            errors.append(f"`{u.name}` diupload di kolom **Income**, tapi isinya file **Order**. Pindahkan ke kolom Order.")
        elif u.kind != "income":
            errors.append(f"`{u.name}` bukan export Income Shopee (tidak ada sheet `Penghasilan`).")

    for u in orders + incomes:
        if u.period is None:
            errors.append(f"Bulan tidak terbaca dari nama file `{u.name}`. Ganti nama file agar memuat "
                          "**satu** nama bulan, misalnya `Order Juli 2026.xlsx` / `Income Juli 2026.xlsx`.")

    # Hanya file yang ada di kolom yang benar dan bulannya terbaca yang dicek duplikat & pasangannya.
    orders = [u for u in orders if u.kind == "order" and u.period]
    incomes = [u for u in incomes if u.kind == "income" and u.period]
    for jenis, group in (("Order", orders), ("Income", incomes)):
        done: list[sm.Period] = []
        for a in group:
            if any(a.period.matches(p) for p in done):
                continue
            done.append(a.period)
            same = [b.name for b in group if a.period.matches(b.period)]
            if len(same) > 1:
                errors.append(f"Ada {len(same)} file {jenis} untuk bulan **{a.period.label}**: "
                              + ", ".join(f"`{n}`" for n in same) + ". Upload satu file per bulan.")

    rows = []
    for o in orders:
        inc = pair_income(o, incomes)
        if inc is None:
            errors.append(f"File Order **{o.period.label}** (`{o.name}`) belum ada pasangannya. "
                          f"Upload juga file **Income {o.period.label}**.")
        rows.append({"Bulan": o.period.label, "File Order": o.name, "File Income": inc.name if inc else "❌ belum ada",
                     "_k": (o.period.year or 0, o.period.month)})
    for i in incomes:
        if not any(o.period.matches(i.period) for o in orders):
            notes.append(f"Income **{i.period.label}** tidak punya file Order bulan yang sama. File ini tetap dipakai "
                         "untuk pesanan bulan sebelumnya yang dananya baru dilepas di bulan itu.")
            rows.append({"Bulan": i.period.label, "File Order": "-", "File Income": i.name,
                         "_k": (i.period.year or 0, i.period.month)})
    pairs = pd.DataFrame(rows)
    if len(pairs):
        pairs = pairs.sort_values("_k").drop(columns="_k").reset_index(drop=True)
    return errors, notes, pairs


def process(orders_up: list[Upload], incomes_up: list[Upload]) -> tuple[sm.MergeResult, list[Output], dict]:
    """Gabungkan file dan buat file Excel hasil (semua di memori)."""
    orders = [sm.read_order(io.BytesIO(u.data), u.name) for u in orders_up]
    incomes = [sm.read_income(io.BytesIO(u.data), u.name) for u in incomes_up]
    result = sm.merge(orders, incomes)
    report = sales_report.build_report(result.gabungan)

    outputs: list[Output] = []
    buf = io.BytesIO()
    sm.to_excel(result, buf)
    outputs.append(Output("Semua periode", "Gabungan Shopee Semua Periode.xlsx", buf.getvalue()))

    if len(orders_up) > 1:
        for u in sorted(orders_up, key=lambda x: (x.period.year or 0, x.period.month) if x.period else (9999, 0)):
            inc = pair_income(u, incomes_up)
            label = u.period.label if u.period else u.name.rsplit(".", 1)[0]
            part = sm.split_by_order_file(result, u.name, inc.name if inc else None)
            buf = io.BytesIO()
            sm.to_excel(part, buf)
            outputs.append(Output(label, f"Gabungan Shopee {label}.xlsx", buf.getvalue()))
    return result, outputs, report


# --------------------------------------------------------------------------- UI
st.title("🧾 Shopee: Gabung Order + Income")
st.caption("File hanya diproses di memori selama sesi ini dan tidak disimpan di server. "
           "Tutup tab atau klik ✕ di file untuk menghapusnya.")

with st.expander("Cara pakai", expanded=False):
    st.markdown(
        """
1. Di Shopee Seller Centre, download:
   - **Pesanan Saya → Export** → file `Order.all....xlsx`
   - **Keuangan → Penghasilan Saya → Export (sudah dilepas)** → file `Income....xlsx`
2. Upload file **Order** di kolom kiri dan file **Income** di kolom kanan (boleh beberapa bulan).
3. **Nama file harus memuat bulan**, misalnya `Order.all Shopee JULI.xlsx` / `Income Juli 2026.xlsx`
   (nama asli Shopee `Order.all.20260601_20260630.xlsx` juga terbaca).
   Setiap file Order bulan X wajib punya file Income bulan X.
4. Klik **Proses**, cek ringkasannya, lalu download hasilnya.

**Catatan aturan gabung**
- Kunci: **No. Pesanan**. 1 baris hasil = 1 baris produk dari file Order.
- Biaya Income diambil **per produk** (baris *Sku* di Laporan Penghasilan).
- Dana pesanan bulan X bisa dilepas di bulan X+1, jadi sebaiknya upload juga Income bulan berikutnya
  agar pesanan akhir bulan ikut tercocokkan.
- Kolom Income yang hanya ada di level pesanan diisi di baris produk pertama saja (agar SUM tidak dobel).
        """
    )

col_order, col_income = st.columns(2)
with col_order:
    order_files = st.file_uploader("📦 File Order (Pesanan Saya → Export)", type=["xlsx"],
                                   accept_multiple_files=True, key="up_order")
with col_income:
    income_files = st.file_uploader("💰 File Income (Penghasilan Saya → sudah dilepas)", type=["xlsx"],
                                    accept_multiple_files=True, key="up_income")


def to_uploads(files, expected: str) -> list[Upload]:
    out = []
    for f in files or []:
        data = f.getvalue()
        out.append(Upload(f.name, data, expected, sm.detect_file_type(io.BytesIO(data)), sm.parse_period(f.name)))
    return out


order_ups = to_uploads(order_files, "order")
income_ups = to_uploads(income_files, "income")

# Hapus hasil lama jika daftar file berubah.
signature = (tuple((u.name, len(u.data)) for u in order_ups), tuple((u.name, len(u.data)) for u in income_ups))
if st.session_state.get("signature") != signature:
    st.session_state.pop("hasil", None)
    st.session_state["signature"] = signature

if not order_ups and not income_ups:
    st.info("Belum ada file. Upload file Order dan file Income untuk bulan yang sama.")
    st.stop()

errors, notes, pairs = validate(order_ups, income_ups)
if len(pairs):
    st.markdown("**Pasangan file per bulan** (bulan dibaca dari nama file)")
    show_df(pairs)
for n in notes:
    st.info(n)
if not order_ups:
    errors.append("Belum ada file **Order**.")
if not income_ups:
    errors.append("Belum ada file **Income**.")
if errors:
    st.error("Perbaiki dulu sebelum diproses:\n\n" + "\n".join(f"- {e}" for e in errors))
    st.stop()

if st.button("Proses", type="primary"):
    with st.spinner("Menggabungkan file..."):
        try:
            st.session_state["hasil"] = process(order_ups, income_ups)
        except Exception as e:  # tampilkan pesan, jangan cetak data ke log
            st.session_state.pop("hasil", None)
            st.error(f"Gagal memproses file: {type(e).__name__}: {e}")

if "hasil" not in st.session_state:
    st.stop()

result, outputs, report = st.session_state["hasil"]
gab = result.gabungan
pesanan = gab.drop_duplicates(sm.KEY)

st.divider()
st.subheader("Hasil")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Baris produk", f"{len(gab):,}")
c2.metric("Pesanan", f"{len(pesanan):,}")
c3.metric("Pesanan ada di Income", f"{(pesanan['Status Match Income'] != 'Belum ada di Income').sum():,}")
c4.metric("Belum ada di Income", f"{(pesanan['Status Match Income'] == 'Belum ada di Income').sum():,}")

belum = pesanan[pesanan["Status Match Income"] == "Belum ada di Income"]
belum_non_batal = belum[belum["Status Pesanan"] != "Batal"]
if len(belum_non_batal):
    st.warning(f"{len(belum_non_batal):,} pesanan **non-Batal** belum ada di Income. "
               "Biasanya dananya belum dilepas, atau Income bulan berikutnya belum diupload.")

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
st.markdown("**Download**")
main, per_period = outputs[0], outputs[1:]
st.download_button(f"⬇️ {main.filename}", data=main.data, file_name=main.filename,
                   mime=XLSX, type="primary", key="dl_all")
if per_period:
    st.caption("Atau per periode (berdasarkan file Order):")
    cols = st.columns(min(len(per_period), 4))
    for i, o in enumerate(per_period):
        cols[i % len(cols)].download_button(f"⬇️ {o.label}", data=o.data, file_name=o.filename,
                                            mime=XLSX, key=f"dl_{i}", width="stretch")

tab0, tab1, tab2, tab3, tab4 = st.tabs(["📊 Report Sales", "Ringkasan", "Status match", "Preview data", "Income tanpa Order"])
with tab0:
    det = report["detail"]
    tot = report["by_produk"].iloc[-1] if len(report["by_produk"]) else None
    if tot is None or det.empty:
        st.info("Belum ada pesanan yang dananya sudah dilepas (ada di Income).")
    else:
        def rp(v: float) -> str:
            if abs(v) >= 1e9:
                return "Rp " + f"{v / 1e9:.2f}".replace(".", ",") + " M"
            if abs(v) >= 1e6:
                return "Rp " + f"{v / 1e6:.1f}".replace(".", ",") + " jt"
            return "Rp " + f"{v:,.0f}".replace(",", ".")

        def full(v: float) -> str:
            return "Rp " + f"{v:,.0f}".replace(",", ".")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Qty Produk (pcs)", f"{tot['Qty Produk']:,.0f}".replace(",", "."))
        m2.metric("Value Gross", rp(tot["Value Gross"]), help=full(tot["Value Gross"]))
        pct = f"{tot['% Total Diskon'] * 100:.1f}".replace(".", ",")
        m3.metric(f"Total Diskon ({pct}%)", rp(tot["Total Diskon"]), help=full(tot["Total Diskon"]) + " dari Value Gross")
        m4.metric("Value Nett", rp(tot["Value Nett"]), help=full(tot["Value Nett"]))
        st.caption("Hanya pesanan yang dananya sudah dilepas (ada di Income). Value Gross = Harga Awal × Jumlah; "
                   "Value Nett = dana diterima (Total Penghasilan); Total Diskon = Gross − Nett.")
        st.markdown("**Per Produk**")
        show_df(fmt_report(report["by_produk"]))
        st.markdown("**Per Bulan & Produk** (bulan pesanan dibuat)")
        show_df(fmt_report(report["by_bulan"]))
        cek = report["mapping"][report["mapping"]["Status"] != "OK"]
        if len(cek):
            st.warning(f"{len(cek)} nama produk perlu dicek pemetaannya (lihat sheet 'Mapping Produk'):")
            show_df(cek[["Nama Produk (Shopee)", "Harga Awal", "Komposisi", "Status"]])
with tab1:
    show_df(result.ringkasan)
with tab2:
    show_df(gab.groupby(["Sumber File Order", "Status Match Income"]).size().rename("Baris").reset_index())
with tab3:
    st.caption(f"Menampilkan 200 baris pertama dari {len(gab):,} baris, {gab.shape[1]} kolom.")
    show_df(gab.head(200))
with tab4:
    orphan = result.income_tanpa_order
    n = orphan[sm.KEY].nunique() if len(orphan) else 0
    st.caption(f"{n:,} pesanan ada di Income tapi file Order-nya tidak diupload (misalnya pesanan bulan sebelumnya).")
    if n:
        show_df(orphan.head(200))

with st.expander("Keterangan aturan gabung"):
    for c in result.catatan:
        st.markdown(f"- {c}")
