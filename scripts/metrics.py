import os
import pandas as pd

D = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data') + os.sep
t = pd.read_csv(D + 'penjualan_vinjimart_okt2026.csv', encoding='utf-8-sig', parse_dates=['tanggal'], dtype={'nota': str})
b = pd.read_csv(D + 'master_barang_vinjimart_2026-11-16.csv', encoding='utf-8-sig', dtype={'kode': str, 'kategori': str, 'satuan': str}, keep_default_na=False, parse_dates=['tgl_masuk'])
for c in ['harga_beli', 'harga_jual', 'stok_awal', 'stok_terjual', 'stok_sisa', 'stok_gudang']:
    b[c] = pd.to_numeric(b[c])
# kolom kasir sudah dianonimkan di CSV: dua akun kasir -> 'Kasir', akun admin -> 'Admin'
t['kasir_anon'] = t.kasir
REF = pd.Timestamp('2026-11-16')

M = {}
# ---- transaksi
M['nota'] = len(t)
M['omzet'] = t.total.sum()
M['modal'] = t.modal.sum()
M['laba'] = t.pendapatan.sum()
M['margin'] = M['laba'] / M['omzet'] * 100
M['hari'] = t.tanggal.nunique()
M['nota_hari'] = M['nota'] / M['hari']
M['omzet_hari'] = M['omzet'] / M['hari']
M['aov'] = t.total.mean()
M['median'] = t.total.median()
d = t.groupby('tanggal').agg(n=('nota', 'size'), omzet=('total', 'sum'), laba=('pendapatan', 'sum'))
M['daily'] = d
M['n_min'], M['n_max'] = d.n.min(), d.n.max()
M['n_min_tgl'], M['n_max_tgl'] = d.n.idxmin(), d.n.idxmax()
M['o_min'], M['o_max'] = d.omzet.min(), d.omzet.max()
M['o_min_tgl'], M['o_max_tgl'] = d.omzet.idxmin(), d.omzet.idxmax()
HARI = {'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu', 'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'}
t['hari'] = t.tanggal.dt.day_name().map(HARI)
nh = d.reset_index().assign(hari=lambda x: x.tanggal.dt.day_name().map(HARI)).groupby('hari').agg(jml_hari=('n', 'size'), nota=('n', 'sum'), omzet=('omzet', 'sum'), laba=('laba', 'sum'))
nh['nota_per_hari'] = nh.nota / nh.jml_hari
nh['omzet_per_hari'] = nh.omzet / nh.jml_hari
M['dow'] = nh.reindex(['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'])
k = t.groupby('kasir_anon').agg(n=('nota', 'size'), omzet=('total', 'sum'), laba=('pendapatan', 'sum'), aov=('total', 'mean'), hari=('tanggal', 'nunique'))
M['kasir'] = k
M['rugi'] = int((t.pendapatan < 0).sum())
M['nol'] = int((t.pendapatan == 0).sum())
bins = [0, 5000, 10000, 25000, 50000, 10**9]
lab = ['≤ Rp 5.000', 'Rp 5.001–10.000', 'Rp 10.001–25.000', 'Rp 25.001–50.000', '> Rp 50.000']
c = pd.cut(t.total, bins=bins, labels=lab)
M['bucket'] = pd.DataFrame({'nota': c.value_counts().reindex(lab), 'omzet': t.groupby(c, observed=False).total.sum().reindex(lab)})
M['hpp_hari'] = M['modal'] / M['hari']

# ---- barang
M['brg'] = len(b)
M['kat_kosong'] = int((b.kategori == '').sum())
M['kat_kosong_pct'] = M['kat_kosong'] / M['brg'] * 100
M['atk'] = int((b.kategori == 'ALAT TULIS KANTOR').sum())
M['satuan_angka'] = int(b.satuan.str.fullmatch(r'\d+').sum())
M['gudang_neg'] = int((b.stok_gudang < 0).sum())
M['gudang_pos'] = int((b.stok_gudang > 0).sum())
M['terjual_gt_awal'] = int((b.stok_terjual > b.stok_awal).sum())
M['stok_tidak_cocok'] = int((b.stok_awal - b.stok_terjual != b.stok_sisa).sum())
M['sisa_neg'] = int((b.stok_sisa < 0).sum())
M['beli0'] = int((b.harga_beli == 0).sum())
M['jual_eq'] = int((b.harga_jual == b.harga_beli).sum())
M['jual_lt'] = int((b.harga_jual < b.harga_beli).sum())
M['kode_dup'] = int(b.kode.duplicated().sum())
M['nama_dup_rows'] = int(b.nama.duplicated(keep=False).sum())
M['kode_pendek'] = int((b.kode.str.strip().str.len() <= 4).sum())
M['tgl_awal'] = b.tgl_masuk.min()
b['umur'] = (REF - b.tgl_masuk).dt.days
dead = b[(b.stok_terjual == 0) & (b.umur >= 90)]
M['dead0'] = int((b.stok_terjual == 0).sum())
M['dead'] = len(dead)
M['dead_modal'] = (dead.stok_sisa.clip(lower=0) * dead.harga_beli).sum()
M['persediaan'] = (b.stok_sisa.clip(lower=0) * b.harga_beli).sum()
M['doi'] = M['persediaan'] / M['hpp_hari']
ok = b[b.harga_beli > 0]
M['margin_brg_median'] = ((ok.harga_jual - ok.harga_beli) / ok.harga_jual).median() * 100
months = (REF - M['tgl_awal']).days / 30.44
tm = M['margin_brg_median'] / 100
low = b[b.harga_jual <= b.harga_beli]
M['harga_gain_thn'] = ((low.harga_beli / (1 - tm) - low.harga_jual).clip(lower=0) * low.stok_terjual).sum() / months * 12
b['omzet_kum'] = b.harga_jual * b.stok_terjual
srt = b.sort_values('omzet_kum', ascending=False)
M['pareto80'] = int((srt.omzet_kum.cumsum() <= 0.8 * b.omzet_kum.sum()).sum()) + 1
M['pareto80_pct'] = M['pareto80'] / M['brg'] * 100
q80 = b.stok_terjual.quantile(0.8)
M['kritis'] = int(((b.stok_terjual >= q80) & (b.stok_sisa <= 2)).sum())
M['top_unit'] = b.nlargest(5, 'stok_terjual')[['nama', 'harga_jual', 'stok_terjual']]
M['top_omzet'] = srt.head(5)[['nama', 'harga_jual', 'stok_terjual', 'omzet_kum']]
M['t'] = t; M['b'] = b

# ---- cost-benefit (36 bulan)
W = 18400  # UMP Sumbar 2026 Rp 3.182.955 / 173 jam
C = {
    'pelatihan': 2 * 3 * W,
    'bersih': 10 * W,
    'ekspor': 39 * W,
    'pelihara': 36 * W,
    'hw_b': 4_500_000,
    'jasa_b': 5_000_000,
}
C['A'] = C['pelatihan'] + C['bersih'] + C['ekspor'] + C['pelihara']
C['B'] = C['A'] + C['hw_b'] + C['jasa_b']
Bn = {
    'harga': M['harga_gain_thn'] * 3,
    'dead': 0.5 * M['dead_modal'],
    'stockout': 0.02 * M['laba'] * 36,
    'waktu': 3 * 36 * W,
}
Bn['total'] = sum(Bn.values())
rutin = (C['ekspor'] + C['pelihara']) / 36
awal_A = C['pelatihan'] + C['bersih']
awal_B = awal_A + C['hw_b'] + C['jasa_b']


def res(benefit, cost, awal):
    net = benefit - cost
    pb = awal / (benefit / 36 - rutin) if benefit / 36 > rutin else float('inf')
    return dict(benefit=benefit, cost=cost, net=net, roi=net / cost * 100, payback=pb)


R = {'A': res(Bn['total'], C['A'], awal_A), 'B': res(Bn['total'], C['B'], awal_B)}
SENS = [
    ('Perhitungan utama Skenario A', res(Bn['total'], C['A'], awal_A)),
    ('A tanpa pencairan dead stock', res(Bn['total'] - Bn['dead'], C['A'], awal_A)),
    ('A, penjualan hilang yang dicegah hanya 1%', res(Bn['total'] - Bn['stockout'] / 2, C['A'], awal_A)),
    ('A tanpa penghematan waktu pemilik', res(Bn['total'] - Bn['waktu'], C['A'], awal_A)),
    ('A kondisi terburuk (ketiganya sekaligus)', res(Bn['harga'] + Bn['stockout'] / 2, C['A'], awal_A)),
    ('Perhitungan utama Skenario B', res(Bn['total'], C['B'], awal_B)),
    ('B tanpa jasa pengembang (hanya beli laptop)', res(Bn['total'], C['B'] - C['jasa_b'], awal_B - C['jasa_b'])),
]
M['C'] = C; M['Bn'] = Bn; M['R'] = R; M['SENS'] = SENS; M['rutin'] = rutin; M['awal_A'] = awal_A; M['awal_B'] = awal_B; M['W'] = W
