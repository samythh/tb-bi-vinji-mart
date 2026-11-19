import os, sys, docx
sys.stdout.reconfigure(encoding='utf-8')
from lib import Builder, set_cover, finalize, rp, num, pct
from metrics import M

FIG = sys.argv[1]
OUT = sys.argv[2]
BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'templates', 'template_m1.docx')
PROYEK = 'Implementasi Business Intelligence melalui Dashboard untuk Visualisasi dan Analisis Penjualan, Produk, dan Persediaan pada Vinji Mart'

d = docx.Document(BASE)
set_cover(d, 'Milestone 1 - Justification & Planning (Vinji Mart)', PROYEK, 'Milestone 1 - Justification & Planning Kelompok 8 kelas BI B')
B = Builder(d)
C, Bn, R, W = M['C'], M['Bn'], M['R'], M['W']
k = M['kasir']
tgl = lambda x: f'{x.day} Oktober 2026'

B.h(2, '4.1 Justification')
B.p('Bagian ini menjelaskan alasan proyek Business Intelligence (BI) Vinji Mart perlu dijalankan. Pembahasan mencakup gambaran umum proyek, penilaian business case, serta rincian analisis biaya dan manfaat.')
B.h(3, '4.1.1 Executive Overview')
B.p('Executive Overview merangkum latar belakang kebutuhan bisnis, tujuan penerapan BI, solusi yang diusulkan, manfaat utama, risiko, dan rekomendasi implementasi.')

B.h(4, '4.1.1.1 Latar Belakang Kebutuhan Bisnis')
B.p('Vinji Mart adalah minimarket di Jl. Parak Karakah No. 6, Kubu Dalam Parak Karakah, Kecamatan Padang Timur, Kota Padang, yang melayani kebutuhan harian warga sekitar dan mahasiswa. Selain barang kebutuhan harian, toko juga melayani transfer, pembayaran e-wallet, pulsa, paket internet, dan token listrik. Toko buka pukul 08.00–22.00 (Senin–Sabtu) dan 09.00–22.00 (Minggu), dijalankan pemilik bersama satu karyawan, dan menerima pembayaran tunai, QRIS, serta transfer. Berbeda dengan usaha yang masih mencatat secara manual, Vinji Mart sudah memakai aplikasi kasir berbasis web, DiaTa Kasir, untuk mencatat setiap penjualan dan mencetak nota. Sistem tersebut menyimpan data master ' + num(M['brg']) + ' barang (SKU) dan riwayat penjualan per nota sejak barang pertama terdaftar pada Januari 2025.')
B.p(f'Laporan penjualan DiaTa Kasir periode 1–31 Oktober 2026 mencatat {num(M["nota"])} nota dengan omzet {rp(M["omzet"])}, harga pokok (modal) {rp(M["modal"])}, dan laba kotor {rp(M["laba"])} atau margin {pct(M["margin"])}. Rata-rata terdapat {num(M["nota_hari"],1)} nota per hari dengan nilai belanja rata-rata {rp(M["aov"])} dan median {rp(M["median"])}. Laporan penjualan mencatat transaksi dari dua akun kasir yang dipakai bergantian, serta sesekali dari akun admin. Pada dataset dan dokumen ini, kedua akun kasir digabung menjadi satu peran Kasir, karena analisis tidak ditujukan untuk menilai kinerja perorangan dan nama kasir merupakan data pribadi.')
B.p('Data tersebut sudah digital, tetapi belum diolah untuk mendukung keputusan. DiaTa Kasir hanya menampilkan laporan berupa tabel statis per nota dan daftar barang. Analisis awal terhadap kedua laporan tersebut menemukan masalah berikut.')
B.ol([
    'Pemilik tidak dapat melihat tren omzet dan laba harian, hari paling ramai, atau perbandingan kategori produk tanpa mengolah ulang tabel laporan secara manual.',
    f'Data master barang belum rapi. Sebanyak {num(M["kat_kosong"])} dari {num(M["brg"])} barang ({pct(M["kat_kosong_pct"])}) tidak memiliki kategori, dan kategori yang terisi pun banyak yang keliru. Contohnya, dari {M["atk"]} barang berkategori Alat Tulis Kantor, sebagian besar justru air mineral, makanan ringan, dan rokok. Kolom satuan pada {num(M["satuan_angka"])} barang berisi angka, bukan satuan barang.',
    f'Catatan stok tidak dapat dipercaya. Stok gudang bernilai negatif pada {num(M["gudang_neg"])} barang dan tidak ada satu pun barang dengan stok gudang positif. Selain itu, {num(M["terjual_gt_awal"])} barang tercatat terjual melebihi stok awalnya. Kondisi ini menunjukkan penerimaan barang dari pemasok belum dicatat sebagai barang masuk.',
    f'Modal tertahan pada barang yang tidak laku. Sebanyak {M["dead"]} barang yang sudah terdaftar lebih dari 90 hari belum pernah terjual, dengan nilai modal {rp(M["dead_modal"])}.',
    f'Ada {M["jual_eq"] + M["jual_lt"]} barang yang harga jualnya sama dengan atau di bawah harga beli ({M["jual_eq"]} sama dan {M["jual_lt"]} di bawah), sehingga barang tersebut terjual tanpa laba.',
    f'Pemesanan ulang ke pemasok belum berbasis data. Sebanyak {M["kritis"]} barang yang termasuk 20% barang terlaris hanya memiliki sisa stok 2 unit atau kurang. Angka pada poin 4 dan 6 dihitung dari stok sisa menurut sistem, sehingga diperlakukan sebagai indikasi awal sampai stok opname dilakukan.',
])

B.h(4, '4.1.1.2 Tujuan Penerapan BI')
B.p('Penerapan BI bertujuan mengubah data DiaTa Kasir menjadi informasi yang mudah dibaca pemilik toko, yaitu untuk:')
B.ol([
    'memantau omzet, laba kotor, margin, jumlah nota, dan nilai belanja rata-rata secara harian, mingguan, dan bulanan;',
    'mengetahui barang dan kategori terlaris serta barang yang tidak laku (dead stock) sebagai dasar pemesanan ulang dan obral;',
    'mendeteksi barang dengan harga jual tidak wajar dan barang laris yang stoknya kritis;',
    'mengukur nilai persediaan dan lama perputaran stok;',
    'membandingkan kinerja penjualan antarhari, antarminggu, dan antarbulan.',
])

B.h(4, '4.1.1.3 Solusi yang Diusulkan')
B.p(f'Solusi berupa alur BI sederhana yang seluruhnya memakai perangkat lunak gratis (Gambar 1). Sumber datanya adalah dua laporan asli dari DiaTa Kasir, yaitu laporan penjualan harian per nota ({num(M["nota"])} nota pada Oktober 2026) dan data barang ({num(M["brg"])} barang, diekspor 16 November 2026). Data diolah dengan skrip Python (Pandas) melalui tahap ekstraksi, pembersihan, standardisasi kategori dan satuan, serta validasi. Hasilnya disimpan di area staging berupa CSV bersih, lalu dimuat ke data mart PostgreSQL berskema bintang dan ditampilkan pada dashboard Power BI Desktop empat halaman: Ringkasan Penjualan, Produk dan Kategori, Persediaan, serta Kontrol Harga dan Kualitas Data.')
B.p('Laporan penjualan DiaTa Kasir hanya memuat total per nota, tanpa rincian barang, jam transaksi, dan metode bayar. Karena itu, analisis produk memakai data stok terjual per barang dari data master. Rincian barang per nota belum termasuk dalam laporan yang diekspor, sehingga analisis keranjang belanja dan jam ramai ditempatkan sebagai pengembangan lanjutan. Untuk analisis tren antarbulan, laporan penjualan bulan lain dapat diekspor dengan mengubah rentang tanggal laporan, karena sistem sudah menyimpan data sejak Januari 2025. Selama ekspor tersebut belum tersedia, tren antarbulan memakai data sintetis yang dibangkitkan mengikuti pola data Oktober 2026.')
B.fig(FIG + '/arsitektur.png', 'Gambar 1. Arsitektur BI Vinji Mart')

B.h(4, '4.1.1.4 Manfaat Utama')
B.p('Manfaat utama proyek ini adalah sebagai berikut.')
B.ol([
    f'Modal tertahan pada {M["dead"]} barang tidak laku senilai {rp(M["dead_modal"])} dapat segera terlihat sehingga barang dapat diobral atau dikembalikan ke pemasok.',
    f'Perbaikan harga pada {M["jual_eq"] + M["jual_lt"]} barang bermargin nol atau negatif berpotensi menambah laba sekitar {rp(M["harga_gain_thn"])} per tahun.',
    f'Barang laris yang stoknya kritis ({M["kritis"]} barang) terpantau lebih awal sehingga risiko kehilangan penjualan berkurang.',
    'Pemilik tidak perlu lagi menyalin dan menjumlahkan laporan DiaTa Kasir secara manual untuk melihat tren dan perbandingan.',
    'Data master barang menjadi rapi karena kategori, satuan, dan harga divalidasi saat ETL dan saat input barang baru.',
])
B.p('Rekap tutup kasir, yaitu menghitung uang di laci dan mencocokkannya dengan laporan harian, tetap dilakukan manual dan tidak dihitung sebagai manfaat, karena DiaTa Kasir sudah menyediakan ringkasan harian.')

B.h(4, '4.1.1.5 Risiko')
B.p('Risiko utama proyek ini beserta mitigasinya adalah sebagai berikut.')
B.ol([
    'Data penjualan yang tersedia saat ini baru satu bulan. Mitigasi: mengekspor laporan bulan sebelumnya dari DiaTa Kasir; selama belum tersedia, data Januari 2025–September 2026 dibangkitkan secara sintetis mengikuti pola Oktober 2026 dan diberi penanda sumber data.',
    'Laporan penjualan tidak memuat rincian barang per nota. Mitigasi: memakai stok terjual per barang dari data master dan meminta laporan rinci bila DiaTa Kasir menyediakannya.',
    'Catatan stok tidak akurat selama barang masuk belum dicatat. Mitigasi: analisis persediaan diberi catatan keterbatasan dan toko menerapkan pencatatan barang masuk.',
    'Ekspor data masih manual karena DiaTa Kasir tidak terhubung langsung ke data mart. Mitigasi: menyusun prosedur ekspor mingguan yang singkat dan skrip ETL yang dapat dijalankan ulang.',
    'Laporan DiaTa Kasir memuat nama kasir. Mitigasi: kedua akun kasir digabung menjadi satu peran Kasir sehingga nama tidak tersimpan di dataset dan dashboard.',
])

B.h(4, '4.1.1.6 Rekomendasi Implementasi')
B.p('Proyek direkomendasikan untuk dijalankan secara bertahap dengan Skenario A. Tahap awal berfokus pada dashboard penjualan dan produk karena datanya sudah tersedia. Tahap berikutnya adalah modul persediaan, yang akurasinya meningkat setelah toko mencatat barang masuk di DiaTa Kasir. Bersamaan dengan itu, toko perlu merapikan kategori dan satuan barang, menetapkan aturan harga jual minimal di atas harga beli, dan menunjuk satu penanggung jawab data. Skenario B tidak direkomendasikan pada skala usaha saat ini karena biayanya belum tertutup manfaat dalam 36 bulan (Tabel 4).')

# ---------------- 4.1.2
B.h(3, '4.1.2 Business Case Assessment Report')
B.p('Penilaian business case disusun berdasarkan sembilan area asesmen yang ditunjukkan pada Tabel 1.')
B.cap('Tabel 1. Business Case Assessment Report')
ra, rb = R['A'], R['B']
B.table(['No', 'Assessment Area', 'Deskripsi'], [
    ['1', 'Strategic Goals', '(1) Menjaga dan meningkatkan laba toko melalui pengelolaan harga dan persediaan yang lebih baik. (2) Mengurangi modal yang tertahan pada barang tidak laku. (3) Menjaga ketersediaan barang laris bagi warga sekitar dan mahasiswa. (4) Membiasakan pengambilan keputusan berbasis data DiaTa Kasir.'],
    ['2', 'Objectives of the Proposed BI Application', '(1) Menyediakan dashboard omzet, laba kotor, margin, jumlah nota, dan AOV. (2) Menampilkan barang dan kategori terlaris serta konsentrasi omzet (Pareto). (3) Menampilkan dead stock, nilai persediaan, dan stok kritis barang laris. (4) Mendeteksi barang dengan harga jual tidak wajar dan data master yang tidak lengkap. (5) Membangun data mart PostgreSQL berskema bintang.'],
    ['3', 'Statement of the Business Need (Business Problem / Business Opportunity)', f'Business Problem: data DiaTa Kasir hanya berupa tabel statis; {pct(M["kat_kosong_pct"])} barang tanpa kategori dan banyak kategori keliru; stok gudang negatif pada {num(M["gudang_neg"])} barang; {M["dead"]} barang tidak laku dengan modal {rp(M["dead_modal"])}; {M["jual_eq"] + M["jual_lt"]} barang bermargin nol atau negatif; {M["kritis"]} barang laris dengan stok kritis. Business Opportunity: {M["pareto80"]} barang ({pct(M["pareto80_pct"])}) menyumbang 80% omzet sehingga pengelolaan stok dapat difokuskan; perbaikan harga dan pencairan dead stock langsung menambah kas dan laba.'],
    ['4', 'Current Decision Support Condition', 'Pemilik membaca laporan harian dan daftar barang di DiaTa Kasir dalam bentuk tabel. Tidak ada grafik tren, perbandingan kategori, atau peringatan stok. Pemesanan ulang ke pemasok mengandalkan pengecekan rak secara langsung.'],
    ['5', 'Explanation of How the BI Application Will Satisfy That Need (Proposed BI Solution)', 'Solusi BI dengan biaya perangkat lunak Rp 0: (1) Sumber data asli DiaTa Kasir, yaitu laporan penjualan per nota dan data barang. (2) ETL Python untuk ekstraksi, pembersihan, standardisasi kategori dan satuan, serta validasi harga dan stok. (3) Staging berupa CSV bersih. (4) Data mart PostgreSQL berskema bintang dengan fakta penjualan per nota dan snapshot persediaan. (5) Dashboard Power BI empat halaman (Gambar 1).'],
    ['6', 'Ramifications of Not Addressing the Business Need and Not Committing to the Proposed BI Solution', '(1) Modal terus tertahan pada barang tidak laku hingga rusak atau kedaluwarsa. (2) Barang laris kehabisan stok sehingga pelanggan beralih ke toko lain. (3) Barang bermargin nol atau negatif terus terjual tanpa laba. (4) Data master makin tidak rapi sehingga laporan DiaTa Kasir makin sulit dipakai.'],
    ['7', 'Cost–Benefit Analysis Results', f'Horizon 36 bulan. Estimated Cost: software Rp 0; hardware Rp 0 (Skenario A) atau {rp(C["hw_b"])} (Skenario B); pelatihan {rp(C["pelatihan"])}; pemeliharaan {rp(C["pelihara"])}; pembersihan data awal {rp(C["bersih"])}; ekspor data rutin {rp(C["ekspor"])}; jasa pengembang Rp 0 (A) atau {rp(C["jasa_b"])} (B). Estimated Benefit: {rp(Bn["total"])}. Skenario A: net benefit {rp(ra["net"])}, ROI {pct(ra["roi"])}, payback ±{num(ra["payback"],1)} bulan. Skenario B: net benefit {rp(rb["net"])}, ROI {pct(rb["roi"])}, payback ±{num(rb["payback"],1)} bulan (melewati horizon). Rincian pada Tabel 2–5.'],
    ['8', 'Recommendations for Data / Business Process Improvement', '(1) Mencatat setiap penerimaan barang sebagai barang masuk di DiaTa Kasir sebelum dipajang. (2) Mengisi kategori dari daftar kategori baku dan memperbaiki kategori yang keliru. (3) Memakai satuan baku (PCS, PACK, BOTOL, dan sebagainya). (4) Menetapkan aturan harga jual minimal di atas harga beli. (5) Memakai kode barang dari barcode, bukan kode pendek buatan. (6) Menunjuk satu penanggung jawab data dan menjadwalkan ekspor laporan mingguan.'],
    ['9', 'Executive Summary', f'Vinji Mart sudah memakai DiaTa Kasir, tetapi datanya belum diolah untuk keputusan. Data Oktober 2026 menunjukkan omzet {rp(M["omzet"])} dengan margin {pct(M["margin"])}, sementara data barang menunjukkan modal tertahan, harga tidak wajar, dan stok yang tidak tercatat. Solusi BI berupa ETL Python, data mart PostgreSQL, dan dashboard Power BI dengan biaya perangkat lunak Rp 0. Skenario A layak dengan ROI {pct(ra["roi"])}; kelayakannya bergantung pada tindak lanjut pemilik terhadap temuan dashboard.'],
], [600, 1700, 6006])

# ---------------- 4.1.3
B.h(3, '4.1.3 Rincian Analisis Biaya–Manfaat')
B.p('Analisis biaya dan manfaat dihitung untuk horizon 36 bulan dengan dua skenario. Skenario A memakai komputer kasir yang sudah ada dan dikerjakan tim mahasiswa. Skenario B mengasumsikan toko membeli satu laptop tambahan untuk pemilik dan membayar jasa pengembang. Rincian biaya dan manfaat ditunjukkan pada Tabel 2 dan Tabel 3.')
B.cap('Tabel 2. Estimasi Biaya (36 bulan)')
B.table(['Komponen', 'Rincian', 'Skenario A', 'Skenario B'], [
    ['Software', 'Power BI Desktop, Python, Pandas, PostgreSQL, draw.io (gratis)', 'Rp 0', 'Rp 0'],
    ['Hardware', 'A: komputer kasir berbasis Windows yang sudah ada. B: satu laptop entry-level untuk pemilik.', 'Rp 0', rp(C['hw_b'])],
    ['Pengembangan', 'A: dikerjakan tim mahasiswa. B: jasa pengembang lepas lokal ±100 jam × Rp 50.000/jam; sengaja di bawah tarif analis BI Indonesia di platform lepas (US$5–45/jam) agar perbandingan tidak merugikan Skenario B.', 'Rp 0', rp(C['jasa_b'])],
    ['Pelatihan', f'2 orang (pemilik dan 1 karyawan) × 3 jam × {rp(W)}/jam', rp(C['pelatihan']), rp(C['pelatihan'])],
    ['Pembersihan data awal', f'{num(M["brg"])} barang × ±15 detik ≈ 10 jam × {rp(W)}/jam', rp(C['bersih']), rp(C['bersih'])],
    ['Ekspor dan refresh data', f'15 menit/minggu × 156 minggu = 39 jam × {rp(W)}/jam', rp(C['ekspor']), rp(C['ekspor'])],
    ['Pemeliharaan', f'1 jam/bulan × 36 bulan × {rp(W)}/jam', rp(C['pelihara']), rp(C['pelihara'])],
    ['**Total Estimasi Biaya**', '', '**' + rp(C['A']) + '**', '**' + rp(C['B']) + '**'],
], [1500, 4006, 1400, 1400])
B.cap('Tabel 3. Estimasi Manfaat Finansial (36 bulan)')
B.table(['Sumber Manfaat', 'Dasar Perhitungan', 'Skenario A', 'Skenario B'], [
    ['Perbaikan harga barang bermargin nol/negatif', f'{M["jual_eq"] + M["jual_lt"]} barang dinaikkan ke margin median {pct(M["margin_brg_median"])}; volume mengikuti stok terjual sejak Januari 2025 yang disetahunkan: {rp(M["harga_gain_thn"])}/tahun × 3 tahun', rp(Bn['harga']), rp(Bn['harga'])],
    ['Pencairan modal dead stock', f'{M["dead"]} barang tidak laku ≥ 90 hari bernilai modal {rp(M["dead_modal"])}; 50% dapat diretur ke distributor (umumnya berlaku untuk barang rusak atau mendekati kedaluwarsa) atau dijual dengan harga obral (sekali)', rp(Bn['dead']), rp(Bn['dead'])],
    ['Pencegahan kehabisan stok barang laris', f'Asumsi 2% penjualan hilang karena stok kosong dapat dicegah: 2% × laba kotor {rp(M["laba"])}/bulan × 36 bulan', rp(Bn['stockout']), rp(Bn['stockout'])],
    ['Penghematan waktu analisis pemilik', f'Pemilik memerlukan ±3 jam/bulan (±45 menit/minggu) untuk membaca dan menjumlahkan laporan; pekerjaan ini digantikan dashboard: 3 × 36 × {rp(W)}/jam', rp(Bn['waktu']), rp(Bn['waktu'])],
    ['**Total Estimasi Manfaat**', '', '**' + rp(Bn['total']) + '**', '**' + rp(Bn['total']) + '**'],
], [1700, 4006, 1300, 1300])
B.p('Manfaat yang belum dapat dinilai dengan uang dicantumkan sebagai intangible benefits pada Tabel 6.')
B.cap('Tabel 4. Ringkasan Hasil Analisis')
B.table(['Indikator', 'Rumus', 'Skenario A', 'Skenario B'], [
    ['Net Benefit', 'Manfaat − Biaya', rp(ra['net']), rp(rb['net'])],
    ['ROI', '(Manfaat − Biaya) / Biaya × 100%', pct(ra['roi']), pct(rb['roi'])],
    ['Payback Period', 'Biaya awal / (manfaat bulanan − biaya rutin bulanan)', f'±{num(ra["payback"],1)} bulan', f'±{num(rb["payback"],1)} bulan (melewati horizon)'],
], [1500, 3206, 1800, 1800])
B.p(f'Dasar perhitungan: nilai waktu staf {rp(W)}/jam berasal dari UMP Sumatera Barat 2026 sebesar Rp 3.182.955 per bulan dibagi 173 jam kerja, dibulatkan. Biaya awal Skenario A adalah pelatihan dan pembersihan data ({rp(M["awal_A"])}), sedangkan Skenario B ditambah laptop dan jasa pengembang ({rp(M["awal_B"])}). Biaya rutin bulanan adalah ekspor data dan pemeliharaan ({rp(M["rutin"])}). Manfaat bulanan rata-rata adalah {rp(Bn["total"] / 36)}. Data stok terjual merupakan angka kumulatif sejak barang terdaftar sehingga disetahunkan terhadap rentang Januari 2025–November 2026.')
B.cap('Tabel 5. Analisis Sensitivitas')
B.table(['Skenario', 'Manfaat (Rp)', 'Biaya (Rp)', 'Net Benefit (Rp)', 'ROI', 'Payback'],
        [[n, num(r['benefit']), num(r['cost']), ('−' if r['net'] < 0 else '') + num(abs(r['net'])), pct(r['roi']), (f'±{num(r["payback"],1)} bulan' + (' (melewati horizon)' if r['payback'] > 36 else ''))] for n, r in M['SENS']],
        [2450, 1150, 1250, 1200, 950, 1306])
worst = M['SENS'][4][1]
B.p(f'Tabel 5 menunjukkan Skenario A tetap layak walaupun ketiga asumsi utama dibuat paling pesimistis (ROI {pct(worst["roi"])}, payback ±{num(worst["payback"],1)} bulan), karena biayanya kecil. Skenario B baru layak bila toko tidak membayar jasa pengembang. Karena itu, keputusan membeli perangkat tambahan sebaiknya ditunda sampai dashboard terbukti dipakai.')
B.cap('Tabel 6. Intangible Benefits')
B.table(['Manfaat', 'Penjelasan'], [
    ['Keputusan stok lebih terarah', 'Pemesanan ulang berdasarkan barang yang benar-benar laku, bukan perkiraan.'],
    ['Data master lebih rapi', 'Kategori, satuan, dan kode barang menjadi seragam sehingga laporan DiaTa Kasir sendiri ikut lebih berguna.'],
    ['Pengawasan operasional', 'Perbandingan penjualan antarhari dan antarminggu membantu pemilik mengatur jadwal jaga toko dan waktu belanja ulang.'],
    ['Dasar pengembangan lanjutan', 'Data mart dapat diperluas untuk analisis keranjang belanja dan jam ramai bila laporan rinci tersedia.'],
], [2500, 5806])

# ---------------- 4.2
B.h(2, '4.2 Planning')
B.p('Tahap Planning bertujuan mengevaluasi kesiapan Vinji Mart, baik teknis maupun nonteknis, serta menyusun rencana proyek BI.')
B.h(3, '4.2.1 Infrastructure Evaluation')
B.p('Evaluasi dilakukan terhadap infrastruktur teknis, yaitu perangkat keras, penyimpanan, DBMS, ETL, dan dashboard, serta infrastruktur nonteknis, yaitu pengelolaan data, kualitas data, tata kelola, proses bisnis, dan kesiapan pengguna.')
B.h(4, '4.2.1.1 Technical Infrastructure Evaluation')
B.cap('Tabel 7. Technical Infrastructure Evaluation Report')
B.table(['Area', 'Kondisi Saat Ini', 'Kebutuhan BI', 'Gap', 'Rekomendasi'], [
    ['Hardware', 'Satu komputer kasir berbasis Windows untuk menjalankan DiaTa Kasir dan mencetak nota. Pengolahan data memakai laptop tim.', 'Laptop/PC Windows 64-bit, RAM minimal 8 GB untuk Power BI Desktop.', 'Perangkat kasir dipakai melayani pembeli sehingga tidak ideal untuk mengolah data saat jam buka.', 'Skenario A: dashboard dibuka di komputer kasir di luar jam sibuk. Skenario B: satu laptop tambahan untuk pemilik.'],
    ['Storage', 'Data tersimpan di server DiaTa Kasir (berbasis web); ekspor berupa file .xls dan cetakan PDF.', 'Penyimpanan terstruktur dengan cadangan; kebutuhan kecil (kurang dari 100 MB).', 'Belum ada arsip ekspor dan cadangan di sisi toko.', 'Folder raw / staging / sql / dashboard dengan cadangan ke Google Drive.'],
    ['DBMS', 'Belum ada database analitik; data hanya di aplikasi DiaTa Kasir.', 'PostgreSQL sebagai data mart berskema bintang.', 'Belum ada database analitik.', 'Memasang PostgreSQL di laptop tim; skrip SQL didokumentasikan.'],
    ['ETL Tools', 'Tidak ada. Laporan penjualan hanya dapat dicetak ke PDF dan data barang diekspor sebagai tabel HTML berekstensi .xls.', 'ETL yang dapat diulang: ekstraksi, pembersihan, standardisasi, validasi, dan load.', 'Belum ada proses ETL; format ekspor tidak langsung siap olah.', 'Skrip Python (Pandas) untuk membaca kedua format ekspor dan memuat data ke PostgreSQL.'],
    ['Dashboard, Reporting and Analytics', 'Laporan tabel di DiaTa Kasir tanpa grafik.', 'Dashboard interaktif dengan KPI, filter, dan tren.', 'Belum ada visualisasi.', 'Power BI Desktop (gratis); refresh manual setelah ETL dijalankan.'],
    ['Backup dan Keamanan', 'Akun DiaTa Kasir dipakai kasir dan admin; laporan memuat nama kasir.', 'Akses dashboard terbatas dan data pribadi disamarkan.', 'Belum ada kebijakan akses untuk data analitik.', 'Password pada perangkat dan file; nama kasir tidak disimpan di dataset.'],
], [1150, 1900, 1700, 1600, 1956])
B.h(4, '4.2.1.2 Nontechnical Infrastructure Evaluation')
B.cap('Tabel 8. Nontechnical Infrastructure Evaluation Report')
B.table(['Area', 'Kondisi Saat Ini', 'Kebutuhan BI', 'Gap', 'Rekomendasi'], [
    ['Data Ownership and Management', 'Data barang diinput melalui akun admin, tetapi belum ada aturan baku pengisian kategori, satuan, dan harga.', 'Ada penanggung jawab data (PIC).', 'Tanggung jawab kerapian data belum jelas.', 'Admin menjadi PIC data; pemilik sebagai penyetuju harga dan kategori.'],
    ['Data Quality', f'{pct(M["kat_kosong_pct"])} barang tanpa kategori; {num(M["satuan_angka"])} satuan berisi angka; {num(M["gudang_neg"])} stok gudang negatif; {M["jual_eq"] + M["jual_lt"]} barang bermargin nol/negatif.', 'Data master lengkap, valid, dan konsisten.', 'Kualitas data master rendah.', 'Pembersihan pada ETL dan perbaikan langsung di DiaTa Kasir (lihat Milestone 2).'],
    ['Data Governance', 'Belum ada aturan penamaan barang, daftar kategori, dan satuan baku.', 'SOP input barang dan daftar kategori baku.', 'Belum ada standar.', 'Menyusun daftar kategori, satuan, dan aturan harga satu halaman.'],
    ['Business Process Readiness', 'Penjualan sudah tercatat rapi di DiaTa Kasir, tetapi barang masuk diduga belum dicatat.', 'Barang masuk dicatat sebelum dipajang.', 'Stok sistem tidak sesuai dengan stok fisik.', 'Menerapkan pencatatan barang masuk dan stok opname awal.'],
    ['User Readiness', 'Pemilik dan kasir terbiasa memakai DiaTa Kasir, tetapi belum pernah memakai dashboard BI.', 'Pengguna mampu membaca KPI dan memakai filter.', 'Kesiapan pengguna rendah.', 'Pelatihan 3 jam dan panduan satu halaman.'],
], [1300, 2000, 1500, 1500, 2006])

B.h(3, '4.2.2 Project Planning')
B.p('Rencana proyek meliputi ruang lingkup, kebutuhan dashboard dan laporan, kebutuhan data historis, keamanan, pengguna, dan hak akses, sebagaimana ditunjukkan pada Tabel 9 dan Tabel 10.')
B.cap('Tabel 9. Project Planning Report')
B.table(['No', 'Project Planning Area', 'Kondisi / Perencanaan Proyek', 'Output yang Diharapkan'], [
    ['1', 'Project Scope', f'IN: dashboard penjualan, produk, persediaan, serta kontrol harga dan kualitas data untuk satu toko Vinji Mart; ETL dari laporan penjualan per nota dan data barang DiaTa Kasir; data mart PostgreSQL berskema bintang; dokumentasi dan pelatihan singkat. OUT: pembuatan atau penggantian aplikasi kasir; layanan transfer, e-wallet, pulsa, dan token listrik; akuntansi dan pajak; integrasi otomatis dengan server DiaTa Kasir; analisis jam ramai, metode bayar, dan keranjang belanja selama laporan rinci belum tersedia. Batasan: perangkat lunak gratis dan refresh manual.', 'Batasan proyek yang jelas.'],
    ['2', 'Project Objectives', 'Menyediakan dashboard omzet, laba kotor, margin, dan AOV; menampilkan barang terlaris dan dead stock; menampilkan nilai persediaan dan stok kritis; mendeteksi harga tidak wajar dan data master tidak lengkap.', 'Tujuan proyek yang terukur.'],
    ['3', 'Functional Deliverables', '(1) Skrip ETL Python dan skrip SQL data mart. (2) Dashboard Power BI (.pbix) empat halaman: Ringkasan Penjualan, Produk dan Kategori, Persediaan, serta Kontrol Harga dan Kualitas Data. (3) Dataset bersih (CSV). (4) Dokumentasi Milestone 1–4 dan kamus data. (5) Panduan pengguna dan prosedur ekspor.', 'Paket proyek yang dapat dijalankan.'],
    ['4', 'Historical Requirements', f'Saat ini tersedia laporan penjualan 1–31 Oktober 2026 ({num(M["nota"])} nota) dan snapshot data barang 16 November 2026. Laporan bulan lain dapat diekspor dari DiaTa Kasir dengan mengubah rentang tanggal laporan. Sampai ekspor tersebut tersedia, data Januari 2025–September 2026 dibangkitkan secara sintetis mengikuti pola data Oktober 2026 (jumlah nota per hari, pola hari dalam seminggu, sebaran nilai nota, dan margin), sesuai izin dosen, sehingga analisis MoM dan YoY (Januari–Oktober) dapat dilakukan. Data sintetis diberi penanda dan diganti ekspor asli bila tersedia.', 'Data historis yang cukup untuk analisis tren.'],
    ['5', 'Subject Area', 'Penjualan (nota, omzet, laba kotor), produk (barang, kategori), persediaan (stok, nilai modal), kasir, dan waktu (tanggal, hari).', 'Fokus area bisnis.'],
    ['6', 'Business Questions / Information Requirements', '(1) Berapa omzet, laba kotor, margin, dan AOV per hari dan per bulan? (2) Hari apa paling ramai? (3) Barang dan kategori apa yang paling laku dan menyumbang omzet terbesar? (4) Barang apa yang tidak laku dan berapa modalnya? (5) Barang laris apa yang stoknya kritis? (6) Barang apa yang harganya tidak wajar? (7) Bagaimana tren penjualan antarbulan?', 'Daftar pertanyaan bisnis.'],
    ['7', 'Condition of Source Files and Databases', f'Laporan penjualan: cetakan PDF {num(M["nota"])} nota, 7 kolom; kolom tanggal kosong sehingga tanggal diturunkan dari nomor nota. Data barang: tabel HTML berekstensi .xls, {num(M["brg"])} barang, 14 kolom; {pct(M["kat_kosong_pct"])} kategori kosong dan stok gudang negatif.', 'Laporan kondisi data sumber.'],
    ['8', 'Stakeholder and Information Users', 'Pemilik (pengambil keputusan; omzet, laba, harga, dan persediaan; prioritas tinggi). Admin/PIC data (kualitas data dan ekspor laporan; tinggi). Kasir (stok kritis dan barang laris; sedang).', 'Daftar pengguna dashboard.'],
    ['9', 'Technology and Tools Planning', 'Python 3 (Pandas, SQLAlchemy), PostgreSQL, Power BI Desktop, draw.io, Microsoft Word.', 'Susunan tools.'],
    ['10', 'Project Timeline / Milestone Planning', 'Pengerjaan dibagi menjadi empat milestone (Tabel 11).', 'Jadwal proyek.'],
], [500, 1600, 4506, 1700])
B.cap('Tabel 10. Pengguna Dashboard, Keamanan, dan Hak Akses')
B.table(['Peran', 'Halaman Dashboard', 'Hak Akses', 'Catatan Keamanan'], [
    ['Pemilik', 'Seluruh halaman', 'Lihat penuh', 'Satu-satunya pengguna yang melihat laba kotor, modal, dan nilai persediaan.'],
    ['Admin / PIC data', 'Persediaan; Kontrol Harga dan Kualitas Data', 'Lihat dan menjalankan ekspor/refresh', 'Tidak mengubah data mart secara langsung.'],
    ['Kasir', 'Produk dan Kategori (tanpa modal)', 'Lihat', 'Tidak melihat laba dan harga beli.'],
    ['Tim pengembang (sementara)', 'Model data dan skrip ETL', 'Kelola', 'Data dipakai untuk keperluan akademik; nama kasir tidak disimpan.'],
], [1700, 2400, 1700, 2506])
B.p('Tabel 10 membagi hak akses berdasarkan peran, bukan jumlah orang; karena toko hanya memiliki pemilik dan satu karyawan, peran kasir dan admin/PIC data dapat dirangkap oleh orang yang sama. Pembatasan akses pada fase ini dilakukan melalui pemisahan file dan password perangkat. Row-level security memerlukan Power BI Service berbayar sehingga menjadi pengembangan lanjutan. Pengelolaan data mengacu pada prinsip pelindungan data pribadi (UU No. 27 Tahun 2022).')
B.cap('Tabel 11. Timeline Proyek')
B.table(['Tahapan', 'Minggu', 'Milestone', 'Output Utama'], [
    ['1. Business Case Assessment (Justification)', '1', 'Milestone 1', 'Executive overview, business case, analisis biaya–manfaat'],
    ['2. Project Planning', '1', 'Milestone 1', 'Evaluasi infrastruktur, rencana proyek'],
    ['3. Business Analysis', '1', 'Milestone 2', 'Proses as-is/to-be, kebutuhan informasi, kualitas data, kebutuhan dashboard'],
    ['4. Design', '2', 'Milestone 3', 'Arsitektur BI, model data mart, rancangan ETL'],
    ['5. ETL dan Data Cleaning', '3', 'Milestone 4', 'Skrip ETL, data mart terisi'],
    ['6. Dashboard Development', '4', 'Milestone 4', 'Dashboard Power BI empat halaman'],
    ['7. Insight, Pengujian, dan Evaluasi', '5', 'Milestone 4', 'Pengujian kecocokan data, insight, rekomendasi'],
], [2700, 900, 1300, 3406])
B.p('Minggu 1 dimulai 17 November 2026. Tenggat Milestone 3 dan Milestone 4 belum ditetapkan sehingga jadwal minggu 2–5 masih perkiraan.')

# ---------------- Lampiran
B.h(2, 'LAMPIRAN')
B.p('Lampiran memuat contoh data asli yang digunakan dalam proyek.')
B.h(3, 'Lampiran A. Dataset/Sample Data')
B.p(f'Proyek memakai dua dataset asli dari DiaTa Kasir. Laporan penjualan harian Oktober 2026 berisi {num(M["nota"])} nota dengan total omzet {rp(M["omzet"])}; hasil ekstraksi telah dicocokkan dengan baris TOTAL pada laporan (modal {rp(M["modal"])} dan laba kotor {rp(M["laba"])}). Data barang berisi {num(M["brg"])} barang per 16 November 2026. Tabel A1 dan A2 menampilkan 10 baris pertama dari masing-masing dataset; kedua akun kasir ditampilkan sebagai satu peran Kasir.')
t = M['t']
B.cap('Tabel A1. Contoh Laporan Penjualan Harian DiaTa Kasir (10 nota pertama, Oktober 2026)')
B.table(['No', 'Tanggal', 'Nota', 'Kasir', 'Modal (Rp)', 'Total (Rp)', 'Laba Kotor (Rp)'],
        [[str(r.no), r.tanggal.strftime('%Y-%m-%d'), r.nota, r.kasir_anon, num(r.modal), num(r.total), num(r.pendapatan)] for r in t.head(10).itertuples()],
        [500, 1250, 1500, 1050, 1300, 1300, 1406])
b = M['b']
B.cap('Tabel A2. Contoh Data Barang DiaTa Kasir (10 baris pertama, 16 November 2026; Awal/Laku/Sisa/Gdg = stok awal, terjual, sisa, dan gudang)')
B.table(['Kode', 'Nama Barang', 'Kategori', 'Sat.', 'Beli (Rp)', 'Jual (Rp)', 'Awal', 'Laku', 'Sisa', 'Gdg'],
        [[r.kode.strip(), r.nama.strip(), r.kategori or '(kosong)', r.satuan, num(r.harga_beli), num(r.harga_jual), num(r.stok_awal), num(r.stok_terjual), num(r.stok_sisa), num(r.stok_gudang)] for r in b.head(10).itertuples()],
        [1300, 1900, 1150, 500, 750, 750, 550, 550, 500, 500], sz=16)

finalize(d)
d.save(OUT)
print('saved', OUT)
