"""
Generator data sintetis Harisna Optical (revisi dari generate_data.py).
Tugas Besar Business Intelligence.

Jalankan: python generate_data_revisi.py
Output  : data objek/transaksi_sintetis_5000_revisi.csv
          data objek/laporan_validasi_sintetis.txt

Semua parameter ada di bagian KONFIGURASI. Kolom "Dasar" menjelaskan asalnya:
  [NOTA]    dihitung dari 21 nota riil
  [TOKO]    keterangan toko/kelompok
  [ASUMSI]  asumsi simulasi, belum didukung data
"""

import csv
import os
import random
from datetime import date, timedelta
from statistics import mean, median

# ============================== KONFIGURASI ==============================
SEED = 42
TOTAL_ROWS = 5000                    # [ASUMSI] skenario pertumbuhan; toko riil 30-60 transaksi/bulan [TOKO]
START_DATE = date(2024, 11, 1)
END_DATE = date(2026, 10, 31)
GROWTH_END = 1.10                    # [ASUMSI] volume akhir periode 10% di atas awal periode
HARI_TUTUP = {date(2025, 3, 31), date(2025, 4, 1), date(2026, 3, 20), date(2026, 3, 21)}  # [ASUMSI] Idulfitri
BOBOT_HARI = [0.25, 0.15, 0.20, 0.12, 0.12, 0.10, 0.06]  # Senin-Minggu; [NOTA] diperhalus, buka tiap hari [TOKO]

# Harga 2026 dipakai untuk seluruh periode. [NOTA] untuk 2026, [ASUMSI] untuk sebelumnya.
MIX_LENSA = {  # nama: (harga, bobot) [NOTA]
    "Rountop": (300000, 0.29),
    "Progresif": (430000, 0.29),
    "Progresif Photokromik": (1100000, 0.14),
    "Progresif Blueguard": (1200000, 0.05),
    "Blueray": (280000, 0.14),
    "Bluechromic": (600000, 0.05),
    "Bluegard": (300000, 0.04),
}
LENSA_ADD = {"Rountop", "Progresif", "Progresif Photokromik", "Progresif Blueguard"}
PELUANG_LA = {  # [NOTA] 4 dari 6 Rountop LA; OP power standar, LA power tinggi/stok kosong [TOKO]
    "Rountop": 4 / 6, "Progresif": 1, "Progresif Photokromik": 1, "Progresif Blueguard": 1,
    "Bluechromic": 1, "Blueray": 0, "Bluegard": 0,
}
HARGA_FRAME = [(205000, .34), (210000, .14), (215000, .14), (220000, .10), (195000, .05),
               (200000, .05), (225000, .05), (0, .13)]  # harga sebelum diskon [NOTA]
PELUANG_DISKON_10 = 0.16             # dari transaksi dengan frame [NOTA]
PELUANG_DISKON_20RB = 0.05           # [NOTA]
PELUANG_KONSUL = 0.14                # konsul saja Rp100.000 [NOTA]
PELUANG_OBAT_KONSUL = 0.57           # Rp130.000-231.000 [NOTA]
PELUANG_LUNAS = 6 / 21               # [NOTA]
RASIO_DP = [.40, .41, .47, .47, .47, .50, .51, .63, .64, .65, .65, .66, .73, .81, .91]  # 15 nota DP [NOTA]
AWALAN = {"Ny.": .55, "Tn.": .25, "An.": .20}  # [NOTA] diperhalus; An. tanpa lensa ber-Add
PELUANG_KEMBALI = 0.355              # [ASUMSI] dikalibrasi agar repeat rate sekitar 30%
JEDA_KEMBALI_MIN = 180               # [ASUMSI] hari
MAKS_KUNJUNGAN = 4                   # [ASUMSI]
PELUANG_POWER_KOSONG = 3 / 21        # data medis di buku dokter [NOTA, TOKO]
PELUANG_DP_TIDAK_DIAMBIL = 0.04      # 3-5% pesanan DP [TOKO]
LEAD_OP = (1, 2)                     # hari ambil OP [TOKO]
LEAD_LA_SELESAI = (7, 10)            # lab Padang [TOKO]
LEAD_LA_AMBIL = (7, 14)              # hari ambil LA [TOKO]
# =========================================================================

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_CSV = os.path.join(BASE, "data objek", "transaksi_sintetis_5000_revisi.csv")
OUT_LOG = os.path.join(BASE, "data objek", "laporan_validasi_sintetis.txt")
REAL_CSV = os.path.join(BASE, "data objek", "transaksi_harisna_gabungan.csv")
INISIAL = "ABCDEFGHIJKLMNPRSTUWYZ"
random.seed(SEED)


def bobot_lensa():
    """Bauran lensa per kelompok, sehingga bauran total tetap mengikuti nota."""
    mix = {k: v[1] for k, v in MIX_LENSA.items()}
    tanpa_add = {k: v for k, v in mix.items() if k not in LENSA_ADD}
    jumlah = sum(tanpa_add.values())
    anak = {k: v / jumlah for k, v in tanpa_add.items()}
    p = AWALAN["An."]
    dewasa = {k: (mix[k] - p * anak.get(k, 0)) / (1 - p) for k in mix}
    return anak, dewasa


BOBOT_ANAK, BOBOT_DEWASA = bobot_lensa()


def pilih(bobot):
    return random.choices(list(bobot), weights=list(bobot.values()), k=1)[0]


def nilai_resep(x):
    """Format nota: +150, -075. Nilai 0 ditulis kosong seperti pada nota."""
    return "" if x == 0 else f"{x:+04d}"


class Pelanggan:
    nomor = 0

    def __init__(self):
        Pelanggan.nomor += 1
        self.id = f"P{Pelanggan.nomor:05d}"
        self.awalan = pilih(AWALAN)
        self.nama = f"{self.awalan} " + "".join(c + "." for c in random.choices(INISIAL, k=random.randint(2, 3)))
        self.pd_jauh = min(69, max(59, round(random.gauss(63, 2.5))))
        self.pd_dekat = self.pd_jauh - random.choice([1, 2, 2, 2, 2, 3, 3])
        self.kunjungan = 0
        self.terakhir = None
        self.tambah = random.choice(range(125, 301, 25))
        if self.awalan == "An.":
            self.sph = random.choice(range(-400, 1, 25))
        elif random.random() < 0.69:          # resep ber-Add dominan plus [NOTA]
            self.sph = random.choice(range(25, 301, 25))
        else:
            self.sph = random.choice(range(-300, 1, 25))
        self.dl = random.choice([-50, -25, 0, 25, 50])

    def resep(self, lensa):
        if random.random() < PELUANG_POWER_KOSONG:
            return ""
        geser = random.choice([-25, 0, 0, 25]) if self.kunjungan > 1 else 0
        r, l = self.sph + geser, self.sph + self.dl + geser
        ada_cyl = random.random() < 0.75
        cr = -random.choice(range(25, 201, 25)) if ada_cyl else 0
        cl = -random.choice(range(25, 201, 25)) if ada_cyl else 0
        axis = f" x{random.choice([90, 180])}" if ada_cyl and random.random() < 1 / 3 else ""

        def mata(s, c, label):
            if s == 0 and c == 0:
                return f"{label}/Plano"
            teks = f"{label}/{nilai_resep(s)}"
            return teks + (f" C{c:04d}{axis}" if c else "")

        teks = f"{mata(r, cr, 'R')} / {mata(l, cl, 'L')}"
        return teks + (f" / Add +{self.tambah:03d}" if lensa in LENSA_ADD else "")


def buat_tanggal():
    hari = [START_DATE + timedelta(days=i) for i in range((END_DATE - START_DATE).days + 1)]
    hari = [h for h in hari if h not in HARI_TUTUP]
    n = (END_DATE - START_DATE).days
    bobot = [BOBOT_HARI[h.weekday()] * (1 + (GROWTH_END - 1) * (h - START_DATE).days / n) for h in hari]
    return sorted(random.choices(hari, weights=bobot, k=TOTAL_ROWS))


def buat_baris():
    rows, semua = [], []
    for tgl in buat_tanggal():
        layak = [c for c in semua if c.kunjungan < MAKS_KUNJUNGAN and (tgl - c.terakhir).days >= JEDA_KEMBALI_MIN]
        if layak and random.random() < PELUANG_KEMBALI:
            c = random.choice(layak)
        else:
            c = Pelanggan()
            semua.append(c)
        c.kunjungan += 1
        c.terakhir = tgl

        lensa = pilih(BOBOT_ANAK if c.awalan == "An." else BOBOT_DEWASA)
        harga_lensa = MIX_LENSA[lensa][0]
        jalur = "LA" if random.random() < PELUANG_LA[lensa] else "OP"
        frame = pilih(dict(HARGA_FRAME))

        ket_diskon, potong = "", 0
        u = random.random()
        if frame and u < PELUANG_DISKON_10:
            frame = int(frame * 0.9) // 1000 * 1000  # 205.000 menjadi 184.000 seperti nota
            ket_diskon = "Diskon 10%"
        elif u < PELUANG_DISKON_10 + PELUANG_DISKON_20RB:
            potong = 20000
            ket_diskon = "Diskon Rp 20.000"

        v = random.random()
        biaya, ket_biaya = 0, ""
        if v < PELUANG_KONSUL:
            biaya, ket_biaya = 100000, "Konsul 100000"
        elif v < PELUANG_KONSUL + PELUANG_OBAT_KONSUL:
            biaya = random.randint(130, 231) * 1000
            ket_biaya = f"Obat & Konsul {biaya}"

        total = frame + harga_lensa + biaya - potong
        lunas = random.random() < PELUANG_LUNAS
        if lunas:
            dp, sisa = total, 0
        else:
            rasio = min(.95, max(.35, random.choice(RASIO_DP) + random.uniform(-.015, .015)))
            dp = round(total * rasio / 1000) * 1000
            sisa = total - dp

        if jalur == "OP":
            selesai = tgl + timedelta(days=1)
            ambil = tgl + timedelta(days=random.randint(*LEAD_OP))
        else:
            selesai = tgl + timedelta(days=random.randint(*LEAD_LA_SELESAI))
            ambil = max(selesai, tgl + timedelta(days=random.randint(*LEAD_LA_AMBIL)))
        lupa = (not lunas) and random.random() < PELUANG_DP_TIDAK_DIAMBIL
        tgl_bayar = tgl if lunas else ambil
        if selesai > END_DATE:
            status, selesai, ambil = "Dikerjakan", None, None
            tgl_bayar = tgl if lunas else None
        elif lupa or ambil > END_DATE:
            status, ambil = "Siap Diambil", None
            tgl_bayar = tgl if lunas else None
        else:
            status = "Diambil"

        ket = "; ".join(x for x in ["Lunas" if lunas else "DP", jalur, ket_biaya, ket_diskon,
                                    f"PD {c.pd_jauh}/{c.pd_dekat}"] if x)
        rows.append({
            "tanggal": tgl.isoformat(), "nama": c.nama, "harga_frame": frame, "harga_lensa": harga_lensa,
            "jenis_lensa": lensa, "total": total, "dp": dp, "sisa": sisa, "power_lensa": c.resep(lensa),
            "keterangan": ket, "id_pelanggan": c.id, "status_pesanan": status,
            "tanggal_selesai": selesai.isoformat() if selesai else "",
            "tanggal_ambil": ambil.isoformat() if ambil else "",
            "tanggal_pelunasan": tgl_bayar.isoformat() if tgl_bayar else "",
            "sumber_data": "Sintetis",
        })
    return rows, semua


def validasi(rows, semua):
    with open(REAL_CSV, encoding="utf-8-sig", newline="") as f:
        real = list(csv.DictReader(f))
    baris = []

    def lapor(label, riil, sint):
        baris.append(f"{label:<34}{riil:>16}{sint:>16}")

    def pct(n, d):
        return f"{n / d * 100:.1f}%"

    n = len(rows)
    rt = [int(r["total"]) for r in real]
    st = [r["total"] for r in rows]
    lapor("Ukuran", "Nota riil", "Sintetis")
    lapor("Jumlah baris", len(real), n)
    lapor("Rata-rata total", f"{mean(rt):,.0f}", f"{mean(st):,.0f}")
    lapor("Median total", f"{median(rt):,.0f}", f"{median(st):,.0f}")
    lapor("Status lunas", pct(sum("Lunas" in r["keterangan"] for r in real), len(real)),
          pct(sum(r["sisa"] == 0 for r in rows), n))
    lapor("Tanpa frame", pct(sum(r["harga_frame"] == "0" for r in real), len(real)),
          pct(sum(r["harga_frame"] == 0 for r in rows), n))
    for lensa, target in [("Rountop", 6), ("Progresif", 6), ("Progresif Photokromik", 3), ("Blueray", 3)]:
        lapor(lensa, pct(target, len(real)), pct(sum(r["jenis_lensa"] == lensa for r in rows), n))
    rasio = sorted(r["dp"] / r["total"] for r in rows if r["sisa"] > 0)
    lapor("Median rasio DP", "62,8%", pct(rasio[len(rasio) // 2], 1))
    ulang = sum(c.kunjungan >= 2 for c in semua)
    lapor("Repeat rate (asumsi)", "-", pct(ulang, len(semua)))
    tertunda = sum(r["status_pesanan"] == "Siap Diambil" and r["sisa"] > 0 for r in rows)
    lapor("DP belum diambil", "-", pct(tertunda, sum(r["sisa"] > 0 for r in rows)))
    baris.append("")

    cek = {
        "dp + sisa = total": all(r["dp"] + r["sisa"] == r["total"] for r in rows),
        "Rountop dan Progresif selalu punya Add": all("Add" in r["power_lensa"] for r in rows
                                                      if r["power_lensa"] and r["jenis_lensa"] in LENSA_ADD),
        "Lensa lain tidak punya Add": all("Add" not in r["power_lensa"] for r in rows
                                          if r["jenis_lensa"] not in LENSA_ADD),
        "An. tidak memakai lensa ber-Add": all(r["jenis_lensa"] not in LENSA_ADD for r in rows if r["nama"].startswith("An.")),
        "Progresif selalu LA": all(" LA;" in " " + r["keterangan"].split("; ", 1)[1][:3] + ";" or "; LA" in r["keterangan"]
                                   for r in rows if r["jenis_lensa"].startswith("Progresif")),
        "Tidak ada transaksi di luar periode": all(START_DATE.isoformat() <= r["tanggal"] <= END_DATE.isoformat() for r in rows),
        "Lunas tanpa sisa": all(r["dp"] == r["total"] for r in rows if r["sisa"] == 0),
    }
    for k, v in cek.items():
        baris.append(f"[{'OK' if v else 'GAGAL'}] {k}")
    teks = "\n".join(baris)
    print(teks)
    with open(OUT_LOG, "w", encoding="utf-8") as f:
        f.write(teks + "\n")
    assert all(cek.values()), "Validasi gagal"


def main():
    rows, semua = buat_baris()
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    print(f"Tersimpan: {OUT_CSV}\n")
    validasi(rows, semua)


if __name__ == "__main__":
    main()
