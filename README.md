# Dashboard Anggaran dan Pendapatan Daerah — Kota Kediri

Situs statis (HTML, CSS, JavaScript) dari LRA Kota Kediri 2023–2025. Tidak perlu build.

## Menjalankan di komputer
`python3 -m http.server 8000` di folder ini, lalu buka http://localhost:8000

## Deploy ke Vercel
Unggah semua isi folder ke GitHub, lalu Import di Vercel dengan Framework Preset "Other". Nama proyek huruf kecil, misalnya `dashboard-apbd-kediri`.

## Logo Kota Kediri
Letakkan file logo di `assets/logo.png`. Situs otomatis memakainya (jika tidak ada, tampil monogram "KD"). Untuk memasukkannya juga ke Excel, jalankan ulang `python3 tools/build_xlsx.py` setelah logo ada.

## Isi
- `downloads/` file Excel per tahun: Dashboard, Rincian LRA, Rumus & Perhitungan, Data Grafik, Catatan & Sumber. Semua hasil hitung berupa rumus.
- Menu **Lembar Kerja**: klik/ketuk sel, alamat dan rumus tampil di bilah fx (juga di HP).
- `tools/` skrip pembuat data. Urutan jika data diubah:
  1. `python3 tools/build_xlsx.py`
  2. hitung ulang tiap file dengan `recalc.py` (skill xlsx) agar nilai tersimpan
  3. `python3 tools/gen_sheets_js.py` dan `python3 tools/gen_data_js.py`
