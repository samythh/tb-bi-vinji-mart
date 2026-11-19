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

## Arsip

Folder `arsip/` menyimpan riwayat pengerjaan, dengan nomor folder sesuai urutan kerja kelompok. Objek awalnya Harisna Optical Pariaman, tetapi ditolak dosen karena toko masih mencatat dengan tulisan tangan dan belum memakai POS.

| Folder | Isi |
|---|---|
| `02 Draft Kelompok` | Draf awal Milestone 1–2 (Harisna Optical) |
| `03 Data` | Data sintetis 5.000 transaksi Harisna Optical |
| `06 Hasil Revisi` | Revisi pertama Milestone 1–2 |
| `07 Dari Sheva` | Versi Sheva, generator data sintetis (`generate_data_revisi.py`), data hasilnya, dan laporan validasi |
| `08 Revisi Final` | Revisi final versi Harisna Optical |
| `09 Siap Kumpul` | Versi Harisna Optical yang siap dikumpulkan (docx + pdf) |
| `10 Sheva Finalisasi` | Scan Milestone 2 yang sudah ditandatangani |
| `11 Dari Lutfi` | Draf Milestone 1 setelah objek diganti ke Vinji Mart |

Catatan untuk versi publik:

- Di Milestone 1 versi Harisna Optical, inisial pasien pada tabel sampel transaksi diganti menjadi `Pasien 01`–`Pasien 15`.
- Tidak ikut diunggah:
  - data asli Harisna Optical, karena berisi data pasien;
  - laporan mentah DiaTa Kasir (pdf dan xls), karena berisi nama kasir;
  - scan Milestone 1 yang sudah ditandatangani, karena tabel pasiennya berupa gambar;
  - slide kuliah dan template dosen;
  - buku referensi, karena dilindungi hak cipta.

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
