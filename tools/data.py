import re, subprocess, os
from data25_23 import parse, R25, R23, n
PDF24 = "/mnt/user-data/uploads/6c7e0ef6-LRA_PEMBIAYAAN_2024.pdf"
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdf2024.txt")

# (no, uraian, level, bold, kode CaLK, anak penjumlah / rumus)
ROWS = [
 (1,"PENDAPATAN - LRA",0,"5.1.1",("sum",[2,7,21])),
 (2,"PENDAPATAN ASLI DAERAH (PAD) - LRA",0,"5.1.1.1",("sum",[3,4,5,6])),
 (3,"Pendapatan Pajak Daerah - LRA",1,"5.1.1.1.1",None),
 (4,"Pendapatan Retribusi Daerah - LRA",1,"5.1.1.1.2",None),
 (5,"Pendapatan Hasil Pengelolaan Kekayaan Daerah yang Dipisahkan - LRA",1,"5.1.1.1.3",None),
 (6,"Lain-lain PAD Yang Sah - LRA",1,"5.1.1.1.4",None),
 (7,"PENDAPATAN TRANSFER - LRA",0,"5.1.1.2",("sum",[8,16,18])),
 (8,"Pendapatan Transfer Pemerintah Pusat - LRA",1,"5.1.1.2.1",("sum",[9,12,13])),
 (9,"Dana Bagi Hasil",2,"",("sum",[10,11])),
 (10,"Bagi Hasil Pajak - LRA",3,"",None),
 (11,"Bagi Hasil Bukan Pajak/Sumber Daya Alam - LRA",3,"",None),
 (12,"Dana Alokasi Umum (DAU) - LRA",2,"",None),
 (13,"Dana Alokasi Khusus - LRA",2,"",("sum",[14,15])),
 (14,"DAK Fisik",3,"",None),
 (15,"DAK Non Fisik",3,"",None),
 (16,"Pendapatan Transfer Pemerintah Pusat - Lainnya - LRA",1,"",("sum",[17])),
 (17,"Dana Insentif Daerah",2,"",None),
 (18,"Pendapatan Transfer Pemerintah Daerah Lainnya - LRA",1,"5.1.1.2.2",("sum",[19,20])),
 (19,"Pendapatan Bagi Hasil Pajak - LRA",2,"",None),
 (20,"Bantuan Keuangan",2,"",None),
 (21,"LAIN-LAIN PENDAPATAN DAERAH YANG SAH - LRA",0,"5.1.1.3",("sum",[22,23])),
 (22,"Pendapatan Hibah - LRA",1,"",None),
 (23,"Lain-lain Pendapatan Sesuai dengan Ketentuan Peraturan Perundang-Undangan",1,"",None),
 (24,"BELANJA",0,"5.1.2",("sum",[25,30,37,39])),
 (25,"BELANJA OPERASI",0,"5.1.2.1",("sum",[26,27,28,29])),
 (26,"Belanja Pegawai",1,"5.1.2.1.1",None),
 (27,"Belanja Barang dan Jasa",1,"5.1.2.1.2",None),
 (28,"Belanja Hibah",1,"5.1.2.1.3",None),
 (29,"Belanja Bantuan Sosial",1,"5.1.2.1.4",None),
 (30,"BELANJA MODAL",0,"5.1.2.2",("sum",[31,32,33,34,35,36])),
 (31,"Belanja Modal Tanah",1,"5.1.2.2.1",None),
 (32,"Belanja Modal Peralatan dan Mesin",1,"5.1.2.2.2",None),
 (33,"Belanja Modal Gedung dan Bangunan",1,"5.1.2.2.3",None),
 (34,"Belanja Modal Jalan, Irigasi dan Jaringan",1,"5.1.2.2.4",None),
 (35,"Belanja Modal Aset Tetap Lainnya",1,"5.1.2.2.5",None),
 (36,"Belanja Modal Aset Lainnya",1,"5.1.2.2.6",None),
 (37,"BELANJA TAK TERDUGA",0,"5.1.2.3",("sum",[38])),
 (38,"Belanja Tak Terduga",1,"",None),
 (39,"BELANJA TRANSFER",0,"5.1.2.4",("sum",[40])),
 (40,"Belanja Bantuan Keuangan",1,"",None),
 (41,"SURPLUS / (DEFISIT)",0,"",("diff",1,24)),
 (42,"PEMBIAYAAN",0,"5.1.3",("head",)),
 (43,"PENERIMAAN PEMBIAYAAN",0,"5.1.3.1",("sum",[44,45])),
 (44,"Penggunaan SiLPA",1,"",None),
 (45,"Pencairan Dana Cadangan",1,"",None),
 (46,"PENGELUARAN PEMBIAYAAN",0,"5.1.3.2",("sum",[47,48])),
 (47,"Penyertaan Modal",1,"",None),
 (48,"Pembentukan Dana Cadangan",1,"",None),
 (49,"PEMBIAYAAN NETTO",0,"",("diff",43,46)),
 (50,"SISA LEBIH PEMBIAYAAN ANGGARAN (SILPA)",0,"",("add",41,49)),
]
KEYS = {r[0]: r for r in ROWS}

def load24():
    if not os.path.exists(CACHE):
        open(CACHE,"w").write(subprocess.run(["pdftotext","-layout",PDF24,"-"],capture_output=True,text=True).stdout)
    d = {}
    for ln in open(CACHE).read().splitlines():
        m = re.match(r'\s*(\d{1,2})\s{2,}\S.*?\s+((?:\(?[\d.]+,\d\d\)?)|-)\s+((?:\(?[\d.]+,\d\d\)?)|-)\s+(?:[\d.]+,\d\d|-)\s+((?:\(?[\d.]+,\d\d\)?)|-)\s*$', ln)
        if m: d[int(m.group(1))] = [n(m.group(2)), n(m.group(3)), n(m.group(4))]
    d[50] = [0.0, 338183377781.83, 379664943066.56]
    return d

def tables():
    """Kembalikan {tahun: {no: (anggaran, realisasi, realisasi_tahun_lalu)}} sesuai dokumen sumber masing-masing."""
    return {2023: parse(R23), 2024: load24(), 2025: parse(R25)}

def evaluate(v):
    """Hitung nilai baris turunan dari anak-anaknya (untuk validasi)."""
    out = {}
    for no, u, lv, k, rule in ROWS:
        if rule is None or rule[0] == "head": continue
        if rule[0] == "sum": out[no] = [sum(v[c][j] for c in rule[1]) for j in range(3)]
        elif rule[0] == "diff": out[no] = [v[rule[1]][j] - v[rule[2]][j] for j in range(3)]
        elif rule[0] == "add": out[no] = [v[rule[1]][j] + v[rule[2]][j] for j in range(3)]
    return out
