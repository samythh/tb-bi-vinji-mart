import os, sys, docx
sys.stdout.reconfigure(encoding='utf-8')
from lib import Builder, set_cover, finalize, rp, num, pct
from metrics import M

FIG = sys.argv[1]
OUT = sys.argv[2]
BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates', 'template_m2.docx')
PROYEK = 'Implementasi Business Intelligence melalui Dashboard untuk Visualisasi dan Analisis Penjualan, Produk, dan Persediaan pada Vinji Mart'

d = docx.Document(BASE)
set_cover(d, 'Milestone 2 - Business Analysis (Vinji Mart)', PROYEK, 'Milestone 2 - Business Analysis Kelompok 8 kelas BI B')
B = Builder(d)
k = M['kasir']
dow = M['dow']
low = M['jual_eq'] + M['jual_lt']

B.h(2, '4.3 Business Analysis')
B.p('Tahap ini menganalisis proses bisnis, kebutuhan informasi, sumber dan kualitas data, serta kebutuhan dashboard dan data yang akan digunakan dalam implementasi BI di Vinji Mart. Analisis memakai dua data asli dari aplikasi kasir DiaTa Kasir, yaitu laporan penjualan harian Oktober 2026 dan data barang per 16 November 2026, ditambah keterangan pihak toko.')

# 4.3.1
B.h(3, '4.3.1 Analisis Proses Bisnis')
B.p('Analisis proses bisnis membandingkan proses saat ini (as-is) dengan proses yang diusulkan (to-be) untuk menemukan titik munculnya masalah data dan keputusan.')
B.h(4, '4.3.1.1 Proses Bisnis Saat Ini (As-Is)')
B.p('Tabel 1 dan Gambar 1 menggambarkan alur penjualan dan pengelolaan barang di Vinji Mart saat ini beserta masalah pada setiap langkah, berdasarkan wawancara dengan pemilik dan kasir pada 17 November 2026.')
B.cap('Tabel 1. Proses As-Is Vinji Mart')
B.table(['No', 'Aktivitas', 'Aktor', 'Data / Dokumen', 'Masalah'], [
    ['1', 'Pemasok atau sales mengantar barang; barang dicek lalu langsung dipajang di rak.', 'Pemasok, Pemilik/Admin', 'Nota pemasok', f'Barang masuk tidak dicatat di DiaTa Kasir sehingga stok gudang negatif pada {num(M["gudang_neg"])} barang.'],
    ['2', 'Barang baru didaftarkan di DiaTa Kasir (kode, nama, kategori, satuan, harga beli, harga jual).', 'Admin', 'Data barang DiaTa Kasir', f'Kategori kosong pada {pct(M["kat_kosong_pct"])} barang dan banyak yang keliru; satuan diisi angka; {low} barang berharga jual ≤ harga beli.'],
    ['3', 'Pelanggan memilih barang lalu membawanya ke kasir.', 'Pelanggan', '–', '–'],
    ['4', 'Kasir memindai atau mencari barang di DiaTa Kasir; sistem menghitung total dan mencetak nota.', 'Kasir', 'Nota, laporan penjualan', 'Data tersimpan rapi, tetapi laporan hanya memuat total per nota tanpa rincian barang dan jam.'],
    ['5', 'Pelanggan membayar secara tunai, QRIS, atau transfer.', 'Pelanggan, Kasir', 'Uang laci', 'Metode bayar tidak muncul di laporan penjualan.'],
    ['6', 'Saat tutup toko, kasir menghitung uang di laci dan mencocokkannya dengan laporan harian DiaTa Kasir.', 'Kasir', 'Laporan harian', 'Langkah ini sudah didukung sistem dan tetap manual.'],
    ['7', 'Pemilik membaca laporan harian dan daftar barang dalam bentuk tabel.', 'Pemilik', 'Laporan DiaTa Kasir', 'Tidak ada grafik tren, perbandingan kategori, maupun peringatan stok.'],
    ['8', 'Pemilik memutuskan belanja ulang ke pemasok.', 'Pemilik', 'Pengecekan rak', f'Tidak berbasis data; {M["kritis"]} barang laris berstok kritis dan {M["dead"]} barang tidak laku tidak terpantau.'],
], [450, 2300, 1150, 1300, 3106])
B.fig(FIG + '/asis_tobe.png', 'Gambar 1. Alur Proses As-Is dan To-Be Vinji Mart', width_cm=15.0)

B.h(4, '4.3.1.2 Proses Bisnis yang Diusulkan (To-Be)')
B.p('Tabel 2 menjelaskan proses setelah pencatatan barang dirapikan dan BI diterapkan. Proses melayani pelanggan tidak berubah; yang berubah adalah cara mencatat barang dan cara memakai data untuk keputusan.')
B.cap('Tabel 2. Proses To-Be dengan BI')
B.table(['No', 'Aktivitas', 'Aktor', 'Sistem / Data', 'Perbaikan'], [
    ['1', 'Barang dari pemasok dicatat sebagai barang masuk di DiaTa Kasir sebelum dipajang.', 'Admin', 'DiaTa Kasir', 'Stok sistem sesuai stok fisik; stok gudang tidak lagi negatif.'],
    ['2', 'Barang baru wajib memakai kategori dari daftar baku, satuan baku, dan harga jual di atas harga beli.', 'Admin', 'DiaTa Kasir, daftar kategori', 'Data master lengkap dan valid sejak input.'],
    ['3', 'Pelanggan memilih barang, kasir memindai barang, dan nota tercetak.', 'Pelanggan, Kasir', 'DiaTa Kasir', 'Tidak berubah.'],
    ['4', 'Saat tutup toko, uang laci dihitung dan dicocokkan dengan laporan harian.', 'Kasir', 'Laporan harian', 'Tetap manual.'],
    ['5', 'Setiap minggu admin mengekspor laporan penjualan dan data barang.', 'Admin (PIC data)', 'File ekspor', 'Data analitik selalu terbaru (±15 menit/minggu).'],
    ['6', 'Skrip ETL Python membersihkan data, menyimpan ke staging, dan memuat ke data mart PostgreSQL; Power BI di-refresh.', 'Admin / tim', 'Python, PostgreSQL, Power BI', 'Proses dapat diulang dan hasilnya konsisten.'],
    ['7', 'Pemilik memantau omzet, laba, barang laris, dead stock, stok kritis, dan harga tidak wajar.', 'Pemilik', 'Dashboard', 'Keputusan belanja ulang, obral, dan harga berbasis data.'],
    ['8', 'Pemilik memesan ulang barang laris yang stoknya kritis dan mengobral atau meretur barang tidak laku.', 'Pemilik, Pemasok', 'Daftar dari dashboard', 'Modal tidak tertahan dan stok kosong berkurang.'],
], [450, 2700, 1150, 1500, 2506])

B.h(4, '4.3.1.3 Gap Analysis')
B.p('Tabel 3 merangkum selisih antara proses as-is dan to-be beserta penanganannya.')
B.cap('Tabel 3. Gap Analysis As-Is dan To-Be')
B.table(['Aspek', 'As-Is', 'To-Be', 'Gap', 'Penanganan'], [
    ['Pemantauan penjualan', 'Tabel per nota', 'Grafik omzet, laba, dan nota per hari', 'Belum ada model data dan dashboard', 'ETL, data mart, Power BI'],
    ['Produk terlaris', 'Tidak terlihat', 'Peringkat barang dan kategori, Pareto', 'Kategori kosong/keliru', 'Standardisasi kategori'],
    ['Persediaan', 'Stok gudang negatif', 'Nilai persediaan dan stok kritis yang akurat', 'Barang masuk tidak dicatat', 'Pencatatan barang masuk dan stok opname awal'],
    ['Dead stock', 'Tidak terpantau', 'Daftar barang tidak laku dan modalnya', 'Tidak ada analisis umur barang', 'Flag dead stock di ETL'],
    ['Harga', f'{low} barang tanpa laba', 'Peringatan harga tidak wajar', 'Tidak ada validasi harga', 'Aturan harga dan flag di ETL'],
    ['Tren antarbulan', 'Tidak ada', 'MoM dan YoY', 'Baru satu bulan laporan diekspor', 'Ekspor laporan bulan lain; sementara memakai data sintetis'],
    ['Analisis per barang per nota dan jam', 'Tidak ada', 'Keranjang belanja dan jam ramai', 'Laporan tidak memuat rincian', 'Laporan rinci DiaTa Kasir (pengembangan lanjutan)'],
], [1500, 1500, 1900, 1700, 1706])

# 4.3.2
B.h(3, '4.3.2 Information Requirements Analysis')
B.p('Tahap ini mengidentifikasi kebutuhan informasi berdasarkan pengguna, pertanyaan bisnis, dan KPI yang diperlukan (Tabel 4).')
B.cap('Tabel 4. Analisis Kebutuhan Informasi')
B.table(['No', 'Pengguna', 'Pertanyaan Bisnis', 'Informasi Dibutuhkan', 'KPI'], [
    ['1', 'Pemilik', 'Bagaimana omzet dan laba kotor per hari, minggu, dan bulan?', 'Total, modal, dan tanggal per nota', 'Total Omzet, Laba Kotor, Margin Kotor'],
    ['2', 'Pemilik', 'Hari apa yang paling ramai?', 'Jumlah nota dan omzet per hari dalam seminggu', 'Nota per Hari, Omzet per Hari'],
    ['3', 'Pemilik', 'Berapa nilai belanja pelanggan dan bagaimana sebarannya?', 'Total per nota', 'AOV, Distribusi Nilai Nota'],
    ['4', 'Pemilik', 'Barang dan kategori apa yang paling laku dan menyumbang omzet terbesar?', 'Stok terjual, harga jual, kategori', 'Top Barang, Kontribusi Kategori, Pareto 80%'],
    ['5', 'Pemilik', 'Barang apa yang tidak laku dan berapa modal yang tertahan?', 'Stok terjual, stok sisa, harga beli, tanggal terdaftar', 'Jumlah Dead Stock, Modal Dead Stock'],
    ['6', 'Pemilik, Admin', 'Barang laris apa yang stoknya hampir habis?', 'Stok terjual dan stok sisa', 'Stok Kritis Barang Laris'],
    ['7', 'Pemilik', 'Berapa nilai persediaan dan berapa lama stok habis terjual?', 'Stok sisa, harga beli, modal penjualan harian', 'Nilai Persediaan, Days of Inventory'],
    ['8', 'Pemilik, Admin', 'Barang apa yang harganya tidak wajar?', 'Harga beli dan harga jual', 'Barang Margin ≤ 0'],
    ['9', 'Admin', 'Seberapa lengkap dan benar data master barang?', 'Kategori, satuan, stok gudang', 'Kelengkapan Kategori, Stok Negatif'],
    ['10', 'Pemilik', 'Bagaimana penjualan pada awal, pertengahan, dan akhir bulan?', 'Tanggal dan total per nota', 'Omzet per Minggu'],
    ['11', 'Pemilik', 'Bagaimana tren penjualan antarbulan?', 'Laporan penjualan beberapa bulan', 'Growth MoM, YoY'],
    ['12', 'Pemilik', 'Jam berapa toko paling ramai dan barang apa yang sering dibeli bersamaan?', 'Jam dan rincian barang per nota', 'Jam Ramai, Market Basket'],
], [450, 1100, 2700, 2050, 2006])
B.p('Kebutuhan nomor 1–10 dapat dipenuhi dengan data yang sudah tersedia, dengan catatan nomor 4–7 bergantung pada standardisasi kategori dan akurasi stok (Tabel 7). Kebutuhan nomor 11 memakai data sintetis Januari 2025–September 2026 yang mengikuti pola Oktober 2026 sampai ekspor asli tersedia, sedangkan nomor 12 memerlukan laporan rinci per barang dan jam transaksi sehingga menjadi pengembangan lanjutan.')

# 4.3.3
B.h(3, '4.3.3 Source Data and Quality Analysis')
B.p('Tahap ini mengidentifikasi sumber data, jenis dan format data, jumlah data, serta permasalahan kualitas data beserta penanganannya.')
B.cap('Tabel 5. Identifikasi Sumber Data')
B.table(['No', 'Nama Sumber Data', 'Deskripsi Data', 'Jumlah Data', 'Keterangan', 'Permasalahan', 'Penanganan'], [
    ['1', 'Laporan Penjualan Harian DiaTa Kasir (cetakan PDF)', 'Satu baris per nota: No, Tanggal, Nota, Kasir, Modal, Total Harga, Pendapatan (laba kotor)', f'{num(M["nota"])} baris, 7 kolom (Okt 2026)', 'Fakta penjualan utama; data asli', f'(1) Kolom tanggal kosong. (2) Teks PDF tidak dapat diseleksi karena dicetak sebagai gambar vektor. (3) {M["rugi"]} nota bermodal lebih besar dari total dan {M["nol"]} nota tanpa laba. (4) Tidak ada rincian barang, jam, dan metode bayar. (5) Nama kasir merupakan data pribadi. (6) Hanya satu bulan.', '(1) Tanggal diturunkan dari 6 digit awal nomor nota (yymmdd). (2) Ekstraksi dengan pengenalan bentuk huruf, divalidasi terhadap baris TOTAL. (3) Ditandai sebagai nota rugi. (4) Ekspor laporan rinci bila tersedia. (5) Dua akun kasir digabung menjadi satu peran Kasir; nama tidak disimpan. (6) Ekspor bulan lain; sementara data Januari 2025–September 2026 dibangkitkan secara sintetis mengikuti pola Oktober 2026.'],
    ['2', 'Data Barang DiaTa Kasir (data_barang_2026-11-16.xls)', 'Kode, nama, kategori, satuan, harga beli, harga jual, stok awal, stok terjual, stok sisa, stok gudang, tanggal barang masuk', f'{num(M["brg"])} baris, 14 kolom (16 Nov 2026)', 'Dimensi barang dan snapshot persediaan; data asli', f'(1) File sebenarnya tabel HTML. (2) {pct(M["kat_kosong_pct"])} kategori kosong dan banyak kategori keliru. (3) {num(M["satuan_angka"])} satuan berisi angka. (4) Stok gudang negatif. (5) Persamaan stok tidak konsisten. (6) Harga tidak wajar. (7) Kode ganda.', '(1) Dibaca sebagai HTML. (2) Mapping kategori baku dengan kata kunci lalu diperiksa manual. (3) Satuan baku. (4–5) Diberi catatan keterbatasan dan stok opname. (6) Flag harga. (7) Penanda duplikat.'],
    ['3', 'Wawancara dan observasi toko', 'Alur proses, jam buka, peran staf, cara belanja ulang', 'Pemilik dan kasir Vinji Mart; 17 November 2026', 'Dasar proses as-is dan asumsi biaya–manfaat', 'Bersifat keterangan lisan', 'Dicatat dalam catatan wawancara kelompok; diperbarui bila ada perubahan'],
], [450, 1100, 1150, 1000, 1250, 1700, 1656])

B.p('Tabel 6 merangkum hasil pemeriksaan kualitas data berdasarkan enam dimensi kualitas.')
B.cap('Tabel 6. Profil Kualitas Data')
B.table(['Dimensi', 'Pemeriksaan', 'Laporan Penjualan', 'Data Barang'], [
    ['Completeness', 'Nilai kosong', f'Tanggal kosong pada {num(M["nota"])} baris (100%)', f'Kategori kosong {num(M["kat_kosong"])} ({pct(M["kat_kosong_pct"])}); harga beli 0 pada {M["beli0"]} barang'],
    ['Validity', 'Nilai sesuai aturan', 'Seluruh nomor nota berformat tanggal yang sah', f'Satuan berisi angka {num(M["satuan_angka"])}; kategori keliru (contoh: {M["atk"]} barang Alat Tulis Kantor yang kebanyakan bukan alat tulis); stok sisa negatif {M["sisa_neg"]}; kode pendek tidak baku {M["kode_pendek"]}'],
    ['Consistency', 'Kecocokan antarkolom', f'Modal + laba = total pada seluruh nota; {M["rugi"]} nota rugi', f'Stok awal − terjual ≠ sisa pada {num(M["stok_tidak_cocok"])} barang; terjual > stok awal pada {num(M["terjual_gt_awal"])}; stok gudang negatif {num(M["gudang_neg"])} dan tidak ada yang positif; harga jual = beli {M["jual_eq"]}, < beli {M["jual_lt"]}'],
    ['Uniqueness', 'Duplikat', f'{num(M["nota"])} nomor nota unik', f'{M["kode_dup"]} kode barang dipakai dua barang; {M["nama_dup_rows"]} baris dengan nama barang ganda'],
    ['Accuracy', 'Kecocokan dengan sumber', f'Jumlah hasil ekstraksi sama persis dengan baris TOTAL laporan ({rp(M["omzet"])}; modal {rp(M["modal"])}; laba {rp(M["laba"])})', 'Stok sistem belum sesuai stok fisik; stok opname dijadwalkan sebagai langkah awal implementasi'],
    ['Timeliness', 'Rentang waktu', 'Satu bulan data asli (Oktober 2026); bulan lain sementara memakai data sintetis', 'Snapshot 16 November 2026; tidak ada riwayat stok'],
], [1300, 1400, 2400, 3206])

B.p('Tabel 7 menilai kesesuaian data yang tersedia dengan kebutuhan analisis.')
B.cap('Tabel 7. Evaluasi Kesesuaian Data terhadap Kebutuhan Analisis')
B.table(['Kebutuhan Analisis', 'Data Tersedia', 'Kesesuaian', 'Tindak Lanjut'], [
    ['Omzet, laba kotor, margin, AOV, nota per hari', 'Laporan penjualan', 'Sesuai (data asli)', '–'],
    ['Perbandingan antarhari dan antarminggu', 'Tanggal dari nomor nota', 'Sesuai', '–'],
    ['Barang dan kategori terlaris', 'Stok terjual, kategori', 'Sesuai sebagian: stok terjual bersifat kumulatif dan kategori perlu distandardisasi', 'Standardisasi kategori; laporan per barang per periode'],
    ['Dead stock dan nilai persediaan', 'Stok sisa, harga beli, tanggal masuk', 'Sesuai sebagian: stok belum sesuai fisik', 'Pencatatan barang masuk dan stok opname'],
    ['Harga tidak wajar', 'Harga beli dan harga jual', 'Sesuai', 'Periksa barang dengan satuan beli dan jual berbeda'],
    ['Tren MoM dan YoY', 'Satu bulan data asli', 'Sesuai dengan data sintetis', 'Data sintetis Januari 2025–September 2026; diganti ekspor asli bila tersedia'],
    ['Jam ramai, metode bayar, keranjang belanja', 'Tidak tersedia', 'Belum sesuai', 'Pengembangan lanjutan'],
], [2300, 1900, 2200, 1906])

# 4.3.4
B.h(3, '4.3.4 Dashboard and Data Requirements')
B.p('Tahap ini menentukan KPI, visualisasi, dan data yang dibutuhkan dashboard. Dashboard terdiri atas empat halaman, yaitu Ringkasan Penjualan, Produk dan Kategori, Persediaan, serta Kontrol Harga dan Kualitas Data (Tabel 8).')
B.cap('Tabel 8. Analisis Kebutuhan Informasi dan Dashboard')
B.table(['No', 'Kebutuhan Informasi', 'Data yang Dibutuhkan', 'Jenis Visualisasi', 'Tujuan Analisis', 'Halaman', 'Status'], [
    ['1', 'Omzet, laba kotor, margin, jumlah nota, AOV', 'Fact_Penjualan_Nota', 'KPI Card', 'Memantau kinerja penjualan', 'Ringkasan Penjualan', 'Tersedia'],
    ['2', 'Tren omzet dan laba harian', 'Fact_Penjualan_Nota, Dim_Tanggal', 'Line Chart', 'Melihat naik turun penjualan', 'Ringkasan Penjualan', 'Tersedia'],
    ['3', 'Nota dan omzet per hari dalam seminggu', 'Fact_Penjualan_Nota, Dim_Tanggal', 'Column Chart', 'Mengatur jadwal dan stok', 'Ringkasan Penjualan', 'Tersedia'],
    ['4', 'Sebaran nilai nota', 'Fact_Penjualan_Nota', 'Histogram', 'Memahami pola belanja', 'Ringkasan Penjualan', 'Tersedia'],
    ['5', 'Penjualan per minggu dalam sebulan', 'Fact_Penjualan_Nota, Dim_Tanggal', 'Column Chart', 'Melihat pengaruh awal dan akhir bulan', 'Ringkasan Penjualan', 'Tersedia'],
    ['6', 'Barang terlaris dan Pareto', 'Fact_Persediaan_Snapshot, Dim_Barang', 'Bar Chart, Pareto Chart', 'Fokus stok pada barang utama', 'Produk dan Kategori', 'Tersedia'],
    ['7', 'Kontribusi kategori', 'Fact_Persediaan_Snapshot, Dim_Barang', 'Treemap', 'Membandingkan kategori', 'Produk dan Kategori', 'Setelah standardisasi'],
    ['8', 'Nilai persediaan dan Days of Inventory', 'Fact_Persediaan_Snapshot, Fact_Penjualan_Nota', 'KPI Card', 'Mengukur modal di rak', 'Persediaan', 'Tersedia'],
    ['9', 'Dead stock dan stok kritis', 'Fact_Persediaan_Snapshot, Dim_Barang', 'Tabel dengan slicer', 'Menentukan obral dan belanja ulang', 'Persediaan', 'Tersedia'],
    ['10', 'Barang margin ≤ 0', 'Fact_Persediaan_Snapshot', 'Tabel', 'Memperbaiki harga', 'Kontrol Harga dan Kualitas Data', 'Tersedia'],
    ['11', 'Kelengkapan data master', 'Dim_Barang, Fact_Persediaan_Snapshot', 'Gauge, Tabel', 'Memantau perbaikan data', 'Kontrol Harga dan Kualitas Data', 'Tersedia'],
    ['12', 'Tren MoM dan YoY', 'Fact_Penjualan_Nota beberapa bulan', 'Line Chart', 'Menilai pertumbuhan', 'Ringkasan Penjualan', 'Data sintetis'],
], [500, 1550, 1500, 1050, 1300, 1300, 1106])

B.p('Definisi dan nilai awal (baseline) setiap KPI ditunjukkan pada Tabel 9. Baseline dihitung dari data asli Oktober 2026 dan data barang 16 November 2026.')
B.cap('Tabel 9. Definisi KPI')
tt = M['t']; wk = tt.groupby((tt.tanggal.dt.day - 1) // 7 + 1).agg(omzet=('total', 'sum'), hari=('tanggal', 'nunique')); wk['ph'] = wk.omzet / wk.hari
B.table(['KPI', 'Rumus', 'Baseline', 'Target'], [
    ['Total Omzet', 'SUM(total)', f'{rp(M["omzet"])}/bulan; {rp(M["omzet_hari"])}/hari', '≥ Rp 27,6 juta/bulan (±10% di atas Oktober 2026)'],
    ['Laba Kotor', 'SUM(total − modal)', f'{rp(M["laba"])}/bulan', 'Naik setelah perbaikan harga'],
    ['Margin Kotor', 'Laba kotor / omzet × 100%', pct(M['margin']), '≥ 15%'],
    ['Jumlah Nota', 'COUNT(nota)', f'{num(M["nota"])}/bulan; {num(M["nota_hari"],1)}/hari (min {M["n_min"]}, maks {M["n_max"]})', 'Dipantau'],
    ['AOV', 'Omzet / jumlah nota', f'{rp(M["aov"])} (median {rp(M["median"])})', 'Naik melalui penataan barang'],
    ['Nota per Hari dalam Seminggu', 'Nota / jumlah hari tersebut', f'Tertinggi Selasa {num(dow.loc["Selasa","nota_per_hari"],1)}; terendah Sabtu {num(dow.loc["Sabtu","nota_per_hari"],1)}', 'Dipakai untuk jadwal dan stok'],
    ['Omzet per Minggu', 'SUM(total) per minggu ke-1 s.d. ke-5', f'Rata-rata per hari tertinggi minggu ke-{wk.ph.idxmax()} ({rp(wk.ph.max())}), terendah minggu ke-{wk.ph.idxmin()} ({rp(wk.ph.min())})', 'Dipakai untuk waktu belanja ulang'],
    ['Pareto 80%', 'Jumlah barang penyumbang 80% omzet kumulatif', f'{M["pareto80"]} barang ({pct(M["pareto80_pct"])})', 'Barang tersebut tidak boleh kosong'],
    ['Nilai Persediaan', 'SUM(stok sisa × harga beli)', rp(M['persediaan']), 'Turun tanpa menambah stok kosong'],
    ['Days of Inventory', 'Nilai persediaan / modal penjualan harian', f'±{num(M["doi"])} hari', '≤ 60 hari'],
    ['Dead Stock', 'Barang terjual 0 dan terdaftar ≥ 90 hari', f'{M["dead"]} barang; modal {rp(M["dead_modal"])}', 'Turun 50% dalam 3 bulan'],
    ['Stok Kritis Barang Laris', '20% barang terlaris dengan stok sisa ≤ 2', f'{M["kritis"]} barang', 'Mendekati 0'],
    ['Barang Margin ≤ 0', 'COUNT(harga jual ≤ harga beli)', f'{low} barang', '0'],
    ['Kelengkapan Kategori', 'Barang berkategori / total barang', pct(100 - M['kat_kosong_pct']), '100% dengan kategori baku'],
    ['Stok Gudang Negatif', 'COUNT(stok gudang < 0)', f'{num(M["gudang_neg"])} barang', '0 setelah pencatatan barang masuk'],
    ['Growth MoM', '(omzet bulan ini − bulan lalu) / bulan lalu', 'Dihitung dari data sintetis sampai ekspor asli tersedia', 'Positif'],
], [1700, 2150, 2500, 1956])

B.h(4, '4.3.4.1 Kebutuhan Data Dimensional')
B.p('Data diorganisasi dalam skema bintang dengan dua tabel fakta yang berbagi dimensi tanggal (Gambar 2). Fact_Penjualan_Nota memiliki grain satu nota penjualan dan berasal dari laporan penjualan harian. Fact_Persediaan_Snapshot memiliki grain satu barang pada satu tanggal ekspor data barang, sehingga riwayat persediaan terbentuk setiap kali data barang diekspor. Bila DiaTa Kasir menyediakan laporan rinci, Milestone 3 akan mengevaluasi fakta tambahan dengan grain satu barang per nota.')
B.fig(FIG + '/star_schema.png', 'Gambar 2. Rancangan Awal Star Schema Vinji Mart', width_cm=15.5)
B.cap('Tabel 10. Kamus Data')
B.table(['Tabel', 'Atribut', 'Tipe', 'Deskripsi', 'Aturan Sumber / Transformasi'], [
    ['Fact_Penjualan_Nota', 'no_nota', 'Teks (PK)', 'Nomor nota DiaTa Kasir', 'Kolom Nota; 10–11 digit'],
    ['Fact_Penjualan_Nota', 'id_tanggal', 'Integer (FK)', 'Kunci ke Dim_Tanggal', 'Dari 6 digit awal nota (yymmdd), format yyyymmdd'],
    ['Fact_Penjualan_Nota', 'id_kasir', 'Integer (FK)', 'Kunci ke Dim_Kasir', 'Dari kolom Kasir: akun kasir menjadi Kasir, akun admin menjadi Admin'],
    ['Fact_Penjualan_Nota', 'modal', 'Currency', 'Harga pokok barang dalam nota', 'Kolom Modal; tanda ribuan dihapus'],
    ['Fact_Penjualan_Nota', 'total', 'Currency', 'Total harga nota', 'Kolom Total Harga'],
    ['Fact_Penjualan_Nota', 'laba_kotor', 'Currency', 'Pendapatan menurut DiaTa Kasir', 'total − modal; nilai negatif dipertahankan'],
    ['Fact_Penjualan_Nota', 'jumlah_nota, sumber_data', 'Integer, Teks', 'Penghitung nota; asal data', "Selalu 1; 'asli' atau 'sintetis'"],
    ['Fact_Persediaan_Snapshot', 'id_barang, id_tanggal_snapshot', 'Integer (FK)', 'Kunci ke Dim_Barang dan Dim_Tanggal', 'Tanggal snapshot dari nama file ekspor'],
    ['Fact_Persediaan_Snapshot', 'harga_beli, harga_jual', 'Currency', 'Harga per barang', 'Tanda ribuan dan desimal ",00" dihapus'],
    ['Fact_Persediaan_Snapshot', 'stok_awal, stok_terjual, stok_sisa, stok_gudang', 'Integer', 'Posisi stok menurut sistem', 'Nilai negatif dipertahankan dan diberi flag'],
    ['Fact_Persediaan_Snapshot', 'nilai_persediaan, margin_persen', 'Currency, Decimal', 'Ukuran turunan', 'max(stok_sisa, 0) × harga_beli; (jual − beli) / jual'],
    ['Fact_Persediaan_Snapshot', 'flag_dead_stock, flag_harga_salah, flag_stok_kritis', 'Boolean', 'Penanda analisis', 'Terjual 0 dan umur ≥ 90 hari; jual ≤ beli; 20% terlaris dengan sisa ≤ 2'],
    ['Dim_Barang', 'id_barang, kode_barang, nama_barang', 'Integer (PK), Teks', 'Identitas barang', 'Kode dirapikan (spasi dihapus); kode ganda diberi akhiran'],
    ['Dim_Barang', 'kategori', 'Teks', 'Kategori baku', 'Mapping kata kunci nama barang lalu diperiksa manual; kosong menjadi "Belum Dikategorikan"'],
    ['Dim_Barang', 'satuan, tgl_terdaftar', 'Teks, Date', 'Satuan baku dan tanggal masuk', 'Satuan angka diganti satuan baku; kolom Tgl Barang Masuk'],
    ['Dim_Kasir', 'id_kasir, peran', 'Integer (PK), Teks', 'Peran akun pencatat', 'Kasir (gabungan dua akun kasir) atau Admin; nama tidak disimpan'],
    ['Dim_Tanggal', 'id_tanggal, tanggal, hari, nama_hari, minggu_ke, bulan, nama_bulan, kuartal, tahun, is_akhir_pekan', 'Integer (PK), Date, Teks, Boolean', 'Atribut kalender', 'Dibangkitkan dengan Pandas'],
], [1950, 1850, 1050, 1450, 2006])

# Lampiran
B.h(2, 'LAMPIRAN')
B.h(3, 'Lampiran A. Ringkasan Data Penjualan Oktober 2026')
B.p('Tabel A1 menampilkan ringkasan penjualan per hari dalam seminggu dan Tabel A2 menampilkan sebaran nilai nota. Keduanya dihitung dari laporan penjualan harian DiaTa Kasir.')
B.cap('Tabel A1. Penjualan per Hari dalam Seminggu (Oktober 2026)')
B.table(['Hari', 'Jumlah Hari', 'Nota', 'Nota per Hari', 'Omzet (Rp)', 'Omzet per Hari (Rp)', 'Laba Kotor (Rp)'],
        [[h, str(int(r.jml_hari)), num(r.nota), num(r.nota_per_hari, 1), num(r.omzet), num(r.omzet_per_hari), num(r.laba)] for h, r in dow.iterrows()],
        [1000, 900, 900, 1100, 1450, 1500, 1456])
bk = M['bucket']
B.cap('Tabel A2. Sebaran Nilai Nota (Oktober 2026)')
B.table(['Nilai Nota', 'Jumlah Nota', '% Nota', 'Omzet (Rp)', '% Omzet'],
        [[str(i), num(r.nota), pct(r.nota / M['nota'] * 100), num(r.omzet), pct(r.omzet / M['omzet'] * 100)] for i, r in bk.iterrows()],
        [2306, 1400, 1300, 1900, 1400])
B.p(f'Lebih dari separuh nota bernilai Rp 10.000 atau kurang, tetapi nota bernilai Rp 25.001–50.000 menyumbang porsi omzet terbesar. Hari dengan omzet tertinggi adalah {M["o_max_tgl"].day} Oktober 2026 ({rp(M["o_max"])}) dan terendah {M["o_min_tgl"].day} Oktober 2026 ({rp(M["o_min"])}).')

finalize(d)
d.save(OUT)
print('saved', OUT)
