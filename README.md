# Tugas Besar BI — Vinji Mart

Tugas besar mata kuliah Business Intelligence, Kelompok 8 kelas B, Departemen Sistem Informasi FTI Universitas Andalas.

**Judul proyek:** Implementasi Business Intelligence melalui Dashboard untuk Visualisasi dan Analisis Penjualan, Produk, dan Persediaan pada Vinji Mart.

Objeknya Vinji Mart, minimarket di Padang Timur yang memakai aplikasi kasir web DiaTa Kasir. Rancangannya ETL Python/Pandas → CSV staging → PostgreSQL → dashboard Power BI.

## Isi repo

| Folder | Isi |
|---|---|
| `docs/` | Milestone 1 (Justification & Planning) dan Milestone 2 (Business Analysis), docx + pdf, diserahkan 19 November 2026 |
| `data/` | Data yang dipakai dokumen (lihat di bawah) |
| `scripts/` | Skrip yang menghitung angka dan membangun dokumen |
| `templates/` | Cover dan daftar isi yang dipakai sebagai kerangka dokumen |
| `gambar/` | Arsitektur BI, alur as-is/to-be, dan rancangan star schema |

## Data

- `penjualan_vinjimart_okt2026.csv` — laporan penjualan harian DiaTa Kasir 1–31 Oktober 2026: 1.507 nota, total Rp 25.117.600, modal Rp 21.483.265, laba kotor Rp 3.634.335. Kolom: `no, tanggal, nota, kasir, modal, total, pendapatan`. Tanggal diambil dari enam digit awal nomor nota (yymmdd). Laporan ini tidak memuat item per nota, jam, atau metode bayar.
- `master_barang_vinjimart_2026-11-16.csv` — ekspor master barang 16 November 2026: 2.421 barang, termasuk harga beli/jual dan stok.

Kolom `kasir` sudah dianonimkan: dua akun kasir ditulis `Kasir` dan akun admin ditulis `Admin`. Sesuai dokumen, nama kasir tidak disimpan.

Data Januari 2025–September 2026 untuk analisis MoM/YoY belum ada di repo. Rencananya data itu dibangkitkan secara sintetis mengikuti pola Oktober 2026 (dengan izin dosen) di Milestone 3, lalu ditandai lewat kolom `sumber_data` (`asli` / `sintetis`).

## Membangun ulang dokumen

Kebutuhan: Python 3 dengan `pandas`, `python-docx`, `matplotlib`. Untuk pdf dan pembaruan daftar isi perlu Microsoft Word di Windows.

```bash
bash scripts/build.sh      # hasil di out/
```

- `metrics.py` menghitung semua angka (omzet, margin, dead stock, cost-benefit, ROI) dari `data/`.
- `m1.py`, `m2.py` menyusun isi Milestone 1 dan 2 di atas template.
- `figs.py` menggambar diagram.
- `finish.ps1` memperbarui daftar isi dan jumlah halaman lalu mengekspor pdf.

## Kelompok 8

| Nama | NIM |
|---|---|
| Luthfi Harisna Mufti (ketua) | 2411523019 |
| Mikail Samyth Habibillah | 2411523016 |
| Sheva Ramadhan | 2411523020 |
| Muhammad Shiddiq Maihendra | 2411523035 |

Dosen pengampu: Rahmatika Pratama Santi, M.T.
