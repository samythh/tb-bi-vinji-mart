import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

OUT = sys.argv[1]
plt.rcParams['font.family'] = 'Times New Roman'


def box(ax, x, y, w, h, title, lines=(), fc='#ffffff', ec='#333333', tsize=10, lsize=8.5, bold=True):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.02,rounding_size=0.08', fc=fc, ec=ec, lw=1.2))
    if lines:
        ax.text(x + w / 2, y + h - 0.22, title, ha='center', va='top', fontsize=tsize, weight='bold' if bold else 'normal')
        ax.text(x + w / 2, y + h - 0.62, '\n'.join(lines), ha='center', va='top', fontsize=lsize, linespacing=1.35)
    else:
        ax.text(x + w / 2, y + h / 2, title, ha='center', va='center', fontsize=tsize, weight='bold' if bold else 'normal', linespacing=1.3)


def arrow(ax, x1, y1, x2, y2):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1), arrowprops=dict(arrowstyle='-|>', lw=1.2, color='#333333'))


# 1. Arsitektur BI (M1)
fig, ax = plt.subplots(figsize=(10, 3.6))
ax.set_xlim(0, 20); ax.set_ylim(0, 7.2); ax.axis('off')
cols = [(0.2, 'Sumber Data', '#e8f0fb'), (5.2, 'Integrasi Data', '#eaf6ec'), (10.2, 'Data Mart', '#fdf3e3'), (15.2, 'Analitik', '#f3eaf8')]
for x, t, c in cols:
    ax.add_patch(FancyBboxPatch((x, 0.2), 4.6, 6.8, boxstyle='round,pad=0.02,rounding_size=0.1', fc=c, ec='none'))
    ax.text(x + 2.3, 6.75, t, ha='center', va='top', fontsize=11, weight='bold')
box(ax, 0.5, 4.0, 4.0, 1.9, 'Laporan Penjualan\nHarian DiaTa Kasir\n(per nota)', tsize=9)
box(ax, 0.5, 1.2, 4.0, 1.9, 'Data Barang\nDiaTa Kasir\n(.xls, 2.421 barang)', tsize=9)
box(ax, 5.5, 4.0, 4.0, 1.9, 'ETL Python (Pandas)\nekstraksi, pembersihan,\nstandardisasi, validasi', tsize=9)
box(ax, 5.5, 1.2, 4.0, 1.9, 'Staging\n(CSV bersih)', tsize=9)
box(ax, 10.5, 2.0, 4.0, 3.2, 'PostgreSQL\nstar schema\nfakta penjualan nota\n+ snapshot persediaan', tsize=9)
box(ax, 15.5, 4.0, 4.0, 1.9, 'Dashboard\nPower BI Desktop\n(4 halaman)', tsize=9)
box(ax, 15.5, 1.2, 4.0, 1.9, 'Pengguna:\nPemilik, Admin/Kasir', tsize=9)
arrow(ax, 4.5, 4.95, 5.5, 4.95); arrow(ax, 4.5, 2.15, 5.5, 4.4)
arrow(ax, 7.5, 4.0, 7.5, 3.1); arrow(ax, 9.5, 2.15, 10.5, 3.2)
arrow(ax, 14.5, 3.9, 15.5, 4.95); arrow(ax, 17.5, 4.0, 17.5, 3.1)
fig.savefig(OUT + '/arsitektur.png', dpi=200, bbox_inches='tight'); plt.close(fig)


# 2. As-Is dan To-Be (M2)
def flow(ax, x, steps, title, fc):
    ax.text(x + 2.6, 15.0, title, ha='center', fontsize=11, weight='bold')
    y = 14.4
    prev = None
    for s, warn in steps:
        h = 1.15
        box(ax, x, y - h, 5.2, h, s, fc=('#fde2e2' if warn else fc), tsize=8.5, bold=False)
        if prev is not None:
            arrow(ax, x + 2.6, prev, x + 2.6, y)
        prev = y - h
        y -= h + 0.45


fig, ax = plt.subplots(figsize=(10, 7.0))
ax.set_xlim(0, 13); ax.set_ylim(2.4, 16.1); ax.axis('off')
asis = [('Pemasok mengantar barang;\nbarang dicek lalu langsung dipajang', True),
        ('Barang masuk tidak dicatat\n(stok gudang negatif)', True),
        ('Barang baru didaftarkan di DiaTa\n(kategori/satuan sering kosong)', True),
        ('Pelanggan memilih barang\nlalu ke kasir', False),
        ('Kasir memindai barang di DiaTa;\nnota tercetak, data tersimpan', False),
        ('Tutup toko: hitung uang laci,\ncocokkan dengan laporan harian', False),
        ('Pemilik melihat tabel laporan;\nbelanja ulang dari cek rak/ingatan', True)]
tobe = [('Barang masuk dicatat di DiaTa\nsebelum dipajang', False),
        ('Barang baru wajib kategori baku,\nsatuan baku, harga jual > beli', False),
        ('Pelanggan memilih barang\nlalu ke kasir', False),
        ('Kasir memindai barang di DiaTa;\nnota tercetak, data tersimpan', False),
        ('Tutup toko: hitung uang laci\n(tetap manual)', False),
        ('Admin ekspor laporan mingguan;\nETL Python ke data mart', False),
        ('Pemilik memantau dashboard:\nomzet, laba, stok, harga', False)]
flow(ax, 0.3, asis, 'As-Is (saat ini)', '#ffffff')
flow(ax, 7.3, tobe, 'To-Be (usulan)', '#eaf6ec')
ax.text(6.5, 2.6, 'Kotak merah: langkah yang menimbulkan masalah data atau keputusan', ha='center', fontsize=8.5, style='italic')
fig.savefig(OUT + '/asis_tobe.png', dpi=200, bbox_inches='tight'); plt.close(fig)

# 3. Star schema (M2)
fig, ax = plt.subplots(figsize=(10, 4.6))
ax.set_xlim(0, 20); ax.set_ylim(0, 9.2); ax.axis('off')
F = '#fdf3e3'; D = '#e8f0fb'
box(ax, 0.3, 6.7, 5.4, 2.2, 'Dim_Kasir', ['id_kasir (PK)', 'peran (Kasir / Admin)'], fc=D)
box(ax, 7.3, 4.75, 5.4, 4.15, 'Fact_Penjualan_Nota', ['no_nota (PK)', 'id_tanggal (FK)', 'id_kasir (FK)', 'modal', 'total', 'laba_kotor', 'jumlah_nota (=1)', 'sumber_data (asli / sintetis)'], fc=F)
box(ax, 14.3, 5.6, 5.4, 3.3, 'Dim_Tanggal', ['id_tanggal (PK)', 'tanggal, hari, nama_hari', 'minggu_ke, bulan, nama_bulan', 'kuartal, tahun', 'is_akhir_pekan'], fc=D)
box(ax, 7.3, 0.2, 5.4, 4.3, 'Fact_Persediaan_Snapshot', ['id_barang (FK)', 'id_tanggal_snapshot (FK)', 'harga_beli, harga_jual', 'stok_awal, stok_terjual', 'stok_sisa, stok_gudang', 'nilai_persediaan, margin_persen', 'flag_dead_stock, flag_harga_salah', 'flag_stok_kritis'], fc=F)
box(ax, 0.3, 0.9, 5.4, 3.0, 'Dim_Barang', ['id_barang (PK)', 'kode_barang, nama_barang', 'kategori (hasil standardisasi)', 'satuan (baku)', 'tgl_terdaftar'], fc=D)
ax.plot([5.7,7.3],[7.4,7.4],color='#333333',lw=1.2); ax.plot([12.7,14.3],[7.4,7.4],color='#333333',lw=1.2)
ax.plot([5.7,7.3],[2.4,2.4],color='#333333',lw=1.2)
ax.plot([17.0,17.0,12.7],[5.6,2.4,2.4],color='#333333',lw=1.2)
ax.text(13.2, 2.65, 'tanggal snapshot', fontsize=8, style='italic')
fig.savefig(OUT + '/star_schema.png', dpi=200, bbox_inches='tight'); plt.close(fig)
print('ok')
