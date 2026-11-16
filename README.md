# streamlitDFD-ShopeeOutputConvert

Aplikasi Streamlit untuk menggabungkan export Shopee **Order.all** (Pesanan Saya) dengan
**Income** (Keuangan → Penghasilan Saya, sudah dilepas) berdasarkan **No. Pesanan**.

## File

| File | Isi |
|---|---|
| `app.py` | UI Streamlit: upload → proses → download |
| `shopee_merge.py` | Logika baca + gabung + tulis Excel (dipakai `app.py` dan `build_hasil.py`) |
| `sales_report.py` | Mapping produk bundling → produk dasar dan pivot Report Sales |
| `build_hasil.py` | Versi lokal: gabungkan semua file di `ACUAN INPUT/` ke `hasil/` |
| `requirements.txt` | Dependensi untuk Streamlit Community Cloud |

## Menjalankan di lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

Atau tanpa UI (folder `ACUAN INPUT` → `hasil`):

```bash
python build_hasil.py
```

## Upload di aplikasi

- Ada 2 kolom upload terpisah: **File Order** dan **File Income** (boleh beberapa bulan).
- **Bulan dibaca dari nama file**, jadi nama file harus memuat satu nama bulan, misalnya
  `Order.all Shopee JULI.xlsx`, `Income.sudah dilepas.id Agustus.xlsx`, `Order Okt 2026.xlsx`.
  Nama asli Shopee dengan tanggal (`Order.all.20260601_20260630.xlsx`) juga terbaca.
- Setiap file Order bulan X **wajib** punya file Income bulan X. Proses ditolak kalau:
  pasangan Income belum ada, bulan tidak terbaca dari nama file, ada 2 file untuk bulan yang sama,
  atau file diupload di kolom yang salah (dicek dari isi file).
- Income tanpa Order di bulan yang sama tetap boleh (dipakai untuk pesanan bulan sebelumnya yang
  dananya baru dilepas).

## Aturan gabung

- 1 baris hasil = 1 baris produk dari file Order (semua kolom Order dipertahankan).
- Kolom Income diberi awalan `Income | `, diambil dari baris **Sku** (rincian per produk) yang cocok
  dengan Nama Produk. Pesanan 1 produk tanpa baris Sku memakai baris **Order**.
- Income dicari di semua file Income yang diupload, karena dana pesanan bulan X bisa dilepas bulan X+1.
- Kolom angka yang hanya ada di level pesanan (mis. `Jumlah Dibayar Pembeli`, `Seller Fee | ...`)
  hanya diisi di baris produk pertama pesanan agar SUM tidak dobel.
- Kolom `Status Match Income` menjelaskan bagaimana tiap baris dicocokkan.

## Report Sales (`sales_report.py`)

Output Excel diawali sheet **Report Sales** (pivot per Produk dan per Bulan × Produk), lalu
**Data Sales** (detail per produk dasar, bisa dibuat Pivot Table sendiri) dan **Mapping Produk**.

| Kolom | Rumus |
|---|---|
| Qty Produk | (Jumlah − Returned quantity) × isi bundling/dus |
| Value Gross | Harga Awal × Jumlah |
| Diskon Penjual | Value Gross − Income Harga Produk |
| Voucher & Diskon Lain | −(voucher/koin penjual + diskon produk dari Shopee) |
| Biaya Marketplace | −(biaya admin, proses, pembayaran, gratis ongkir XTRA, promo XTRA, AMS, PPh 22, …) |
| Ongkir Ditanggung Penjual | −(ongkir pembeli + ongkir ke jasa kirim + subsidi ongkir Shopee + …) |
| Refund | −(Jumlah Pengembalian Dana ke Pembeli) |
| Penyesuaian Lain | sisa selisih dari Shopee (pesanan refund) |
| Total Diskon | jumlah komponen di atas = Value Gross − Value Nett |
| Value Nett | Income "Total Penghasilan" (dana diterima) |

Tanggal dan Bulan di report memakai **Waktu Pengiriman Diatur** (fallback ke Waktu Pesanan Dibuat bila kosong).
Hanya pesanan yang sudah ada di Income yang dihitung. Bundling dipecah ke produk dasar
(Tongkol 250gr, Tongkol 500gr, Tuna Can 165gr, …) dan nilainya dibagi proporsional terhadap harga standar.
Nama produk baru yang tidak dikenali muncul sebagai "Tidak terpetakan" di sheet Mapping Produk —
tambahkan aturannya di `parse_product()`.

## Deploy ke Streamlit Community Cloud

1. Push folder ini ke GitHub (file `.xlsx` dan folder `hasil/` sudah di-ignore, **jangan** commit data Shopee).
2. Di share.streamlit.io → **Create app** → pilih repo, branch, dan main file
   `streamlitDFD-ShopeeOutputConvert/app.py`.
3. Setelah jalan, buka **Settings → Sharing** dan batasi viewer ke email tertentu,
   karena file Shopee berisi data pembeli.

## Privasi

File upload hanya diproses di memori selama sesi. Aplikasi tidak menulis ke disk,
tidak memakai cache untuk data upload, dan tidak mencetak isi data ke log.
