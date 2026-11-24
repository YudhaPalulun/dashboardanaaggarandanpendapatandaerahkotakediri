import re, subprocess, os
from data25_23 import parse, R25, R23, n
PDF24 = "/mnt/user-data/uploads/6c7e0ef6-LRA_PEMBIAYAAN_2024.pdf"
CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdf2024.txt")

# (no, uraian, level, bold, kode CaLK, anak penjumlah / rumus)
ROWS = [
 (41,"SURPLUS / (DEFISIT) – dari LRA",0,"",None),
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
ROW = {r[0]: 5 + i for i, r in enumerate(ROWS)}   # nomor LRA -> baris Excel di Rincian LRA
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
