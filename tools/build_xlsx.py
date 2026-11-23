"""Membuat 3 file Excel (2023, 2024, 2025) dengan rumus hidup. Urutan: build_xlsx.py -> recalc -> gen_sheets_js.py -> gen_data_js.py"""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XLImage
from data import ROWS, KEYS, tables

OUT = os.environ.get("OUT_DIR", os.path.join(HERE, "..", "downloads"))
F = "Arial"
IND = "22265C"; AMB = "F2A93B"; SUBF = "ECEBF7"; GRYF = "F6F4EF"; MUT = "5B5E7A"
RP = '#,##0.00;(#,##0.00);"-"'; PCT = '0.00%;(0.00%);"-"'; MIL = '#,##0.00'
thin = Side(style="thin", color="C9C8DA"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def font(b=False, c="000000", s=10, i=False): return Font(name=F, bold=b, color=c, size=s, italic=i)
def fill(c): return PatternFill("solid", start_color=c, end_color=c)
def row_of(n): return n + 4                      # nomor sumber -> baris Excel di Rincian LRA
def code(c, size=8):
    c.data_type = "s"; c.font = Font(name="Consolas", size=size, color="2B3A8C")
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

def sum_expr(L, kids):
    rs = [row_of(k) for k in kids]
    if len(rs) >= 3 and rs == list(range(rs[0], rs[0] + len(rs))): return f"=SUM({L}{rs[0]}:{L}{rs[-1]})"
    return "=" + "+".join(f"{L}{r}" for r in rs)

def hdr(ws, r, vals, c0=1):
    for j, v in enumerate(vals):
        c = ws.cell(r, c0 + j, v); c.font = font(True, "FFFFFF"); c.fill = fill(IND)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
    ws.row_dimensions[r].height = 34

def title(ws, text, sub, span):
    for r in (1, 2): ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    ws["A1"] = text; ws["A1"].font = font(True, IND, 13); ws["A1"].alignment = Alignment(vertical="center")
    ws["A2"] = sub; ws["A2"].font = font(False, MUT, 9, True); ws.row_dimensions[1].height = 24

def build(Y, T):
    V = T[Y]                                       # {no: (anggaran, realisasi, realisasi tahun lalu)}
    wb = Workbook(); dash = wb.active; dash.title = "Dashboard"
    lra = wb.create_sheet("Rincian LRA"); rm = wb.create_sheet("Rumus & Perhitungan")
    dg = wb.create_sheet("Data Grafik"); cat = wb.create_sheet("Catatan & Sumber")
    L = lambda n, col: f"'Rincian LRA'!{col}{row_of(n)}"

    # ---------------- Rincian LRA ----------------
    title(lra, f"RINCIAN LAPORAN REALISASI ANGGARAN PENDAPATAN DAN BELANJA DAERAH – PEMERINTAH KOTA KEDIRI, TA {Y}",
          "(dalam rupiah) · Sumber: LRA Pemerintah Kota Kediri. Angka biru = salinan dari dokumen sumber; angka hitam = rumus Excel.", 11)
    hdr(lra, 4, ["No", "CaLK", "Uraian", f"Anggaran {Y}", f"Realisasi {Y}", "Capaian (%)", f"Realisasi {Y-1}",
                 "Selisih thd Tahun Lalu (Rp)", "Perubahan (%)", "Sisa / (Lebih) Anggaran (Rp)", "Porsi (%)"])
    for no, ur, lv, kode, rule in ROWS:
        r = row_of(no); bold = lv == 0
        lra.cell(r, 1, no).font = font(c=MUT); lra.cell(r, 2, kode).font = font(c=MUT, s=9)
        c = lra.cell(r, 3, ur); c.font = font(bold); c.alignment = Alignment(indent=lv, wrap_text=True, vertical="center")
        if rule and rule[0] == "head":
            for cc in range(1, 12): lra.cell(r, cc).fill = fill(GRYF)
            continue
        for col, L_, j in ((4, "D", 0), (5, "E", 1), (7, "G", 2)):
            if rule is None:
                c = lra.cell(r, col, V[no][j]); c.font = font(c="0000FF")
            elif rule[0] == "sum": lra.cell(r, col, sum_expr(L_, rule[1]))
            elif rule[0] == "diff": lra.cell(r, col, f"={L_}{row_of(rule[1])}-{L_}{row_of(rule[2])}")
            elif rule[0] == "add": lra.cell(r, col, f"={L_}{row_of(rule[1])}+{L_}{row_of(rule[2])}")
        lra.cell(r, 6, f"=IF(D{r}=0,0,E{r}/D{r})"); lra.cell(r, 8, f"=E{r}-G{r}")
        lra.cell(r, 9, f"=IF(G{r}=0,0,H{r}/ABS(G{r}))"); lra.cell(r, 10, f"=D{r}-E{r}")
        if no <= 23: lra.cell(r, 11, f"=IF($E${row_of(1)}=0,0,E{r}/$E${row_of(1)})")
        elif no <= 40: lra.cell(r, 11, f"=IF($E${row_of(24)}=0,0,E{r}/$E${row_of(24)})")
    for no, ur, lv, kode, rule in ROWS:
        r = row_of(no)
        for cc in range(1, 12):
            cell = lra.cell(r, cc); cell.border = BOX
            if cc in (4, 5, 7, 8, 10): cell.number_format = RP
            if cc in (6, 9, 11): cell.number_format = PCT
            if cc >= 4 and not (cell.font and cell.font.color and cell.font.color.rgb == "000000FF"): cell.font = font(lv == 0 and rule is not None and rule[0] != "head" or no in (41, 49, 50))
            if lv == 0 and rule and rule[0] != "head" and cc <= 11:
                cell.fill = fill(SUBF)
        if no in (1, 24, 50):
            for cc in range(1, 12): lra.cell(r, cc).fill = fill(IND); lra.cell(r, cc).font = font(True, "FFFFFF")
    # baris 3: rumus tiap kolom (contoh) supaya terbaca walau pratinjau HP tanpa bilah rumus
    lra.cell(3, 3, "Rumus tiap kolom (contoh baris 6 = PAD):").font = font(True, IND, 9)
    ex = row_of(2)
    first_kid = row_of(3); last_kid = row_of(6)
    for col, L_ in ((4, "D"), (5, "E"), (7, "G")):
        code(lra.cell(3, col, f"Input (biru)\nSubtotal: =SUM({L_}{first_kid}:{L_}{last_kid})\nTotal: =" + "+".join(f"{L_}{row_of(k)}" for k in (2, 7, 21))))
    for col, ft in ((6, f"=IF(D{ex}=0,0,E{ex}/D{ex})"), (8, f"=E{ex}-G{ex}"), (9, f"=IF(G{ex}=0,0,H{ex}/ABS(G{ex}))"),
                    (10, f"=D{ex}-E{ex}"), (11, f"=IF($E${row_of(1)}=0,0,E{ex}/$E${row_of(1)})")):
        code(lra.cell(3, col, ft))
    lra.row_dimensions[3].height = 66
    for col, w in zip("ABCDEFGHIJK", [6, 11, 52, 23, 23, 12, 23, 23, 13, 23, 11]): lra.column_dimensions[col].width = w
    lra.freeze_panes = "D5"

    # ---------------- Rumus & Perhitungan ----------------
    title(rm, f"RUMUS DAN PERHITUNGAN – ANGGARAN DAN PENDAPATAN DAERAH KOTA KEDIRI TA {Y}",
          "Setiap hasil dihitung dengan rumus Excel yang mengacu ke sheet Rincian LRA. Ubah angka biru di Rincian LRA, maka semua hasil ikut berubah.", 5)
    hdr(rm, 4, ["No", "Indikator", "Rumus Excel yang dipakai", "Penjelasan", "Hasil"])
    ind = [
     ("A. PENDAPATAN", None),
     ("Capaian Pendapatan (%)", (f"={L(1,'E')} / {L(1,'D')}", "Realisasi pendapatan dibagi anggaran (target) pendapatan.", f"=IF({L(1,'D')}=0,0,{L(1,'E')}/{L(1,'D')})", "0.00%")),
     ("Capaian PAD (%)", (f"={L(2,'E')} / {L(2,'D')}", "Realisasi PAD dibagi target PAD.", f"=IF({L(2,'D')}=0,0,{L(2,'E')}/{L(2,'D')})", "0.00%")),
     ("Capaian Pendapatan Transfer (%)", (f"={L(7,'E')} / {L(7,'D')}", "Realisasi transfer dibagi target transfer.", f"=IF({L(7,'D')}=0,0,{L(7,'E')}/{L(7,'D')})", "0.00%")),
     ("Pendapatan melampaui / (di bawah) target (Rp)", (f"={L(1,'E')} - {L(1,'D')}", "Realisasi dikurangi anggaran pendapatan. Positif berarti melampaui target.", f"={L(1,'E')}-{L(1,'D')}", RP)),
     ("Porsi PAD terhadap Pendapatan (%)", (f"={L(2,'E')} / {L(1,'E')}", "Ukuran kemandirian fiskal: seberapa besar pendapatan berasal dari PAD.", f"=IF({L(1,'E')}=0,0,{L(2,'E')}/{L(1,'E')})", "0.00%")),
     ("Porsi Transfer terhadap Pendapatan (%)", (f"={L(7,'E')} / {L(1,'E')}", "Ukuran ketergantungan pada transfer pusat dan daerah.", f"=IF({L(1,'E')}=0,0,{L(7,'E')}/{L(1,'E')})", "0.00%")),
     ("Pertumbuhan Pendapatan (%)", (f"=({L(1,'E')} - {L(1,'G')}) / {L(1,'G')}", f"Perubahan realisasi pendapatan {Y} terhadap {Y-1}.", f"=IF({L(1,'G')}=0,0,({L(1,'E')}-{L(1,'G')})/{L(1,'G')})", "+0.00%;-0.00%;0.00%")),
     ("B. BELANJA", None),
     ("Capaian Belanja (%)", (f"={L(24,'E')} / {L(24,'D')}", "Realisasi belanja dibagi anggaran belanja (daya serap).", f"=IF({L(24,'D')}=0,0,{L(24,'E')}/{L(24,'D')})", "0.00%")),
     ("Sisa Anggaran Belanja (Rp)", (f"={L(24,'D')} - {L(24,'E')}", "Anggaran belanja yang tidak terserap.", f"={L(24,'D')}-{L(24,'E')}", RP)),
     ("Rasio Belanja terhadap Pendapatan (%)", (f"={L(24,'E')} / {L(1,'E')}", "Di atas 100% berarti belanja melebihi pendapatan (defisit).", f"=IF({L(1,'E')}=0,0,{L(24,'E')}/{L(1,'E')})", "0.00%")),
     ("Porsi Belanja Operasi (%)", (f"={L(25,'E')} / {L(24,'E')}", "Belanja operasi dibagi total belanja.", f"=IF({L(24,'E')}=0,0,{L(25,'E')}/{L(24,'E')})", "0.00%")),
     ("Porsi Belanja Modal (%)", (f"={L(30,'E')} / {L(24,'E')}", "Belanja modal dibagi total belanja.", f"=IF({L(24,'E')}=0,0,{L(30,'E')}/{L(24,'E')})", "0.00%")),
     ("Pertumbuhan Belanja (%)", (f"=({L(24,'E')} - {L(24,'G')}) / {L(24,'G')}", f"Perubahan realisasi belanja {Y} terhadap {Y-1}.", f"=IF({L(24,'G')}=0,0,({L(24,'E')}-{L(24,'G')})/{L(24,'G')})", "+0.00%;-0.00%;0.00%")),
     ("C. SURPLUS/DEFISIT DAN PEMBIAYAAN", None),
     ("Surplus / (Defisit) Realisasi (Rp)", (f"={L(1,'E')} - {L(24,'E')}", "Pendapatan dikurangi belanja.", f"={L(1,'E')}-{L(24,'E')}", RP)),
     ("Pembiayaan Netto Realisasi (Rp)", (f"={L(43,'E')} - {L(46,'E')}", "Penerimaan pembiayaan dikurangi pengeluaran pembiayaan.", f"={L(43,'E')}-{L(46,'E')}", RP)),
     ("SILPA Realisasi (Rp)", (f"={L(41,'E')} + {L(49,'E')}", "Surplus/defisit ditambah pembiayaan netto.", f"={L(41,'E')}+{L(49,'E')}", RP)),
     ("D. PEMERIKSAAN OTOMATIS", None),
     ("Pendapatan − Belanja = Surplus/(Defisit)", ("=IF(ABS(E1 - E24 - E41) < 1, \"OK\", \"PERIKSA\")", "Memastikan selisih pendapatan dan belanja sama dengan baris surplus/defisit.", f"=IF(ABS({L(1,'E')}-{L(24,'E')}-{L(41,'E')})<1,\"OK\",\"PERIKSA\")", "General")),
     ("Surplus/(Defisit) + Pembiayaan Netto = SILPA", ("=IF(ABS(E41 + E49 - E50) < 1, \"OK\", \"PERIKSA\")", "Memastikan SILPA konsisten dengan komponennya.", f"=IF(ABS({L(41,'E')}+{L(49,'E')}-{L(50,'E')})<1,\"OK\",\"PERIKSA\")", "General")),
     ("Porsi PAD + Transfer + Lain-lain = 100%", ("=IF(ABS(K6 + K11 + K25 - 1) < 0.000001, \"OK\", \"PERIKSA\")", "Memastikan tiga porsi pendapatan berjumlah 100%.", f"=IF(ABS({L(2,'K')}+{L(7,'K')}+{L(21,'K')}-1)<0.000001,\"OK\",\"PERIKSA\")", "General")),
    ]
    rr = 5; no = 0
    for name, spec in ind:
        if spec is None:
            rm.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=5)
            c = rm.cell(rr, 1, name); c.font = font(True, IND); 
            for cc in range(1, 6): rm.cell(rr, cc).fill = fill(SUBF); rm.cell(rr, cc).border = BOX
            rr += 1; continue
        no += 1; ft, expl, fx, nf = spec
        rm.cell(rr, 1, no).font = font(c=MUT); rm.cell(rr, 2, name).font = font(True)
        c = rm.cell(rr, 3); c.value = ft; c.data_type = "s"; c.font = Font(name="Consolas", size=9, color="2B3A8C")
        rm.cell(rr, 4, expl).font = font(s=9)
        h = rm.cell(rr, 5, fx); h.number_format = nf; h.font = font(True); h.alignment = Alignment(horizontal="right")
        for cc in range(1, 6): rm.cell(rr, cc).border = BOX; rm.cell(rr, cc).alignment = Alignment(wrap_text=True, vertical="center", horizontal="right" if cc == 5 else None)
        rm.row_dimensions[rr].height = 32; rr += 1
    for col, w in zip("ABCDE", [6, 44, 50, 52, 26]): rm.column_dimensions[col].width = w
    rm.freeze_panes = "A5"
    rm_last = rr - 1

    # ---------------- Data Grafik ----------------
    title(dg, f"DATA GRAFIK – KOTA KEDIRI TA {Y}", "Semua angka di sini adalah rumus yang mengambil data dari Rincian LRA (satuan miliar rupiah). Kolom F–H menampilkan rumusnya.", 8)
    blocks = []                                     # (judul, kolom, [(label, no)], tipe)
    blocks.append(("Pendapatan per kelompok (Rp miliar)", ["Anggaran", "Realisasi", "Capaian (%)"], [("PAD", 2), ("Transfer", 7), ("Lain-lain", 21)], "cap"))
    blocks.append(("Rincian PAD (Rp miliar)", ["Anggaran", "Realisasi", "Capaian (%)"], [("Pajak Daerah", 3), ("Retribusi", 4), ("Hasil Kekayaan Dipisahkan", 5), ("Lain-lain PAD Sah", 6)], "cap"))
    blocks.append(("Komposisi Transfer – Realisasi (Rp miliar)", ["Anggaran", "Realisasi", "Porsi thd Transfer (%)"], [("Dana Bagi Hasil", 9), ("DAU", 12), ("DAK", 13), ("Transfer Pusat Lainnya", 16), ("Transfer Pemda Lainnya", 18)], "por7"))
    blocks.append(("Belanja per jenis (Rp miliar)", ["Anggaran", "Realisasi", "Capaian (%)"], [("Operasi", 25), ("Modal", 30), ("Tak Terduga", 37), ("Transfer", 39)], "cap"))
    blocks.append(("Rincian Belanja Operasi (Rp miliar)", ["Anggaran", "Realisasi", "Capaian (%)"], [("Pegawai", 26), ("Barang dan Jasa", 27), ("Hibah", 28), ("Bantuan Sosial", 29)], "cap"))
    r = 4; gpos = {}
    for ti, cols, items, kind in blocks:
        dg.cell(r, 1, ti).font = font(True, IND); hdr(dg, r + 1, ["Kelompok"] + cols); r += 2; start = r
        for lab, n in items:
            dg.cell(r, 1, lab)
            dg.cell(r, 2, f"={L(n,'D')}/10^9"); dg.cell(r, 3, f"={L(n,'E')}/10^9")
            if kind == "cap": dg.cell(r, 4, f"=IF(B{r}=0,0,C{r}/B{r})")
            else: dg.cell(r, 4, f"=IF({L(7,'E')}=0,0,{L(n,'E')}/{L(7,'E')})")
            for k, ftxt in enumerate((f"={L(n,'D')}/10^9", f"={L(n,'E')}/10^9", f"=IF(B{r}=0,0,C{r}/B{r})" if kind == "cap" else f"=IF({L(7,'E')}=0,0,{L(n,'E')}/{L(7,'E')})")):
                code(dg.cell(r, 6 + k, ftxt), 8); dg.cell(r, 6 + k).alignment = Alignment(horizontal="left", vertical="center")
            for cc in range(1, 5):
                dg.cell(r, cc).border = BOX
                if cc in (2, 3): dg.cell(r, cc).number_format = MIL
                if cc == 4: dg.cell(r, cc).number_format = "0.0%"
                if cc == 1: dg.cell(r, cc).font = font()
            r += 1
        gpos[ti] = (start, r - 1); r += 1
    dg.cell(r, 1, "Alur Surplus/(Defisit) dan SILPA – Realisasi (Rp miliar)").font = font(True, IND); hdr(dg, r + 1, ["Komponen", "Anggaran", "Realisasi"]); r += 2; wstart = r
    for lab, n in (("Pendapatan", 1), ("Belanja", 24), ("Surplus/(Defisit)", 41), ("Pembiayaan Netto", 49), ("SILPA", 50)):
        dg.cell(r, 1, lab); dg.cell(r, 2, f"={L(n,'D')}/10^9"); dg.cell(r, 3, f"={L(n,'E')}/10^9")
        code(dg.cell(r, 6, f"={L(n,'D')}/10^9")); code(dg.cell(r, 7, f"={L(n,'E')}/10^9"))
        for k in (6, 7): dg.cell(r, k).alignment = Alignment(horizontal="left", vertical="center")
        for cc in range(1, 4):
            dg.cell(r, cc).border = BOX
            if cc in (2, 3): dg.cell(r, cc).number_format = MIL
        r += 1
    hdr(dg, 3, ["Kolom A–D: hasil rumus", "", "", ""], 1); dg.merge_cells("A3:D3"); dg.cell(3, 6, "Rumus yang dipakai (teks)").font = font(True, IND, 9)
    for col, w in zip("ABCDEFGH", [36, 18, 18, 22, 3, 34, 34, 40]): dg.column_dimensions[col].width = w

    # ---------------- Dashboard ----------------
    dash.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJKL", [3, 20, 20, 3, 20, 20, 3, 20, 20, 3, 3, 3]): dash.column_dimensions[col].width = w
    dash.merge_cells("A1:L1"); dash["A1"] = f"DASHBOARD ANGGARAN DAN PENDAPATAN DAERAH – KOTA KEDIRI, TA {Y}"
    dash["A1"].font = font(True, "FFFFFF", 15); dash["A1"].fill = fill(IND); dash["A1"].alignment = Alignment(vertical="center", indent=1); dash.row_dimensions[1].height = 50
    dash.merge_cells("A2:L2"); dash["A2"] = "Semua angka pada kartu di bawah adalah rumus yang mengambil data dari sheet Rincian LRA. Rumusnya ditulis di bawah tiap kartu."
    dash["A2"].font = font(False, MUT, 9, True)
    if os.path.exists(os.path.join(HERE, "..", "assets", "logo.png")):
        lg = XLImage(os.path.join(HERE, "..", "assets", "logo.png")); _w, _h = lg.width, lg.height; lg.height = 58; lg.width = round(58 * _w / _h); dash.add_image(lg, "I1")
    kpis = [
        ("Realisasi Pendapatan (Rp miliar)", f"={L(1,'E')}/10^9", MIL, f"={L(1,'E')}/10^9"),
        ("Capaian Pendapatan", f"=IF({L(1,'D')}=0,0,{L(1,'E')}/{L(1,'D')})", "0.00%", "=E_pendapatan / D_pendapatan"),
        ("Porsi PAD terhadap Pendapatan", f"=IF({L(1,'E')}=0,0,{L(2,'E')}/{L(1,'E')})", "0.00%", "=E_PAD / E_pendapatan"),
        ("Realisasi Belanja (Rp miliar)", f"={L(24,'E')}/10^9", MIL, f"={L(24,'E')}/10^9"),
        ("Surplus / (Defisit) (Rp miliar)", f"={L(41,'E')}/10^9", '#,##0.00;(#,##0.00)', f"={L(41,'E')}/10^9"),
        ("SILPA (Rp miliar)", f"={L(50,'E')}/10^9", MIL, f"={L(50,'E')}/10^9"),
    ]
    pos = [(4, 2), (4, 5), (4, 8), (8, 2), (8, 5), (8, 8)]
    for (lab, fx, nf, ftxt), (r0, c0) in zip(kpis, pos):
        dash.merge_cells(start_row=r0, start_column=c0, end_row=r0, end_column=c0 + 1); dash.merge_cells(start_row=r0 + 1, start_column=c0, end_row=r0 + 1, end_column=c0 + 1)
        dash.merge_cells(start_row=r0 + 2, start_column=c0, end_row=r0 + 2, end_column=c0 + 1)
        a = dash.cell(r0, c0, lab); a.font = font(True, "FFFFFF", 9); a.fill = fill(IND); a.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        b = dash.cell(r0 + 1, c0, fx); b.font = font(True, IND, 17); b.number_format = nf; b.fill = fill(SUBF); b.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c = dash.cell(r0 + 2, c0, fx); c.data_type = "s"; c.font = Font(name="Consolas", size=7, color="2B3A8C"); c.fill = fill(GRYF); c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        for rr_, colr in ((r0, IND), (r0 + 1, SUBF), (r0 + 2, GRYF)):
            for cc_ in (c0, c0 + 1): dash.cell(rr_, cc_).fill = fill(colr)
        dash.row_dimensions[r0].height = 20; dash.row_dimensions[r0 + 1].height = 34; dash.row_dimensions[r0 + 2].height = 26
    for rr_ in (3, 7, 11): dash.row_dimensions[rr_].height = 8

    def style(ch, t):
        ch.title = t; ch.roundedCorners = False; ch.width = 15.8; ch.height = 8.6
        if ch.legend is not None: ch.legend.overlay = False; ch.legend.position = "b"
        if ch.title is not None: ch.title.overlay = False
    ps, pe = gpos["Pendapatan per kelompok (Rp miliar)"]
    c1 = BarChart(); c1.type = "col"; c1.add_data(Reference(dg, min_col=2, max_col=3, min_row=ps - 1, max_row=pe), titles_from_data=True); c1.set_categories(Reference(dg, min_col=1, min_row=ps, max_row=pe))
    style(c1, "Pendapatan: Anggaran vs Realisasi (Rp miliar)"); c1.x_axis.delete = False; c1.y_axis.delete = False; c1.y_axis.numFmt = "#,##0"
    c1.series[0].graphicalProperties.solidFill = "B9BCE0"; c1.series[1].graphicalProperties.solidFill = IND
    dash.add_chart(c1, "B12")
    ts, te = gpos["Komposisi Transfer – Realisasi (Rp miliar)"]
    c2 = DoughnutChart(); c2.add_data(Reference(dg, min_col=3, min_row=ts - 1, max_row=te), titles_from_data=True); c2.set_categories(Reference(dg, min_col=1, min_row=ts, max_row=te))
    style(c2, "Komposisi Pendapatan Transfer – Realisasi"); c2.dataLabels = DataLabelList(); c2.dataLabels.showPercent = True; c2.dataLabels.showVal = False; c2.dataLabels.showCatName = False; c2.dataLabels.showSerName = False; c2.dataLabels.showLeaderLines = False
    dash.add_chart(c2, "B30")
    bs, be = gpos["Belanja per jenis (Rp miliar)"]
    c3 = BarChart(); c3.type = "bar"; c3.add_data(Reference(dg, min_col=2, max_col=3, min_row=bs - 1, max_row=be), titles_from_data=True); c3.set_categories(Reference(dg, min_col=1, min_row=bs, max_row=be))
    style(c3, "Belanja per Jenis: Anggaran vs Realisasi (Rp miliar)"); c3.x_axis.delete = False; c3.y_axis.delete = False; c3.y_axis.numFmt = "#,##0"
    c3.series[0].graphicalProperties.solidFill = "F6D9A6"; c3.series[1].graphicalProperties.solidFill = AMB
    dash.add_chart(c3, "B48")

    # ---------------- Catatan & Sumber ----------------
    title(cat, "CATATAN DAN SUMBER DATA", f"Dokumen: Laporan Realisasi Anggaran (LRA) Pemerintah Kota Kediri, TA {Y} dan {Y-1}.", 2)
    notes = [
        ("Sumber", f"LRA Pemerintah Kota Kediri untuk tahun yang berakhir sampai dengan 31 Desember {Y} dan {Y-1}."),
        ("Cakupan", "Pendapatan (PAD, Transfer, Lain-lain), Belanja (Operasi, Modal, Tak Terduga, Transfer), Surplus/Defisit, dan Pembiayaan hingga SILPA. Seluruh baris nomor 1 sampai 50 pada LRA."),
        ("Angka biru", "Salinan angka dari LRA (anggaran, realisasi tahun berjalan, realisasi tahun lalu) untuk baris rincian."),
        ("Angka hitam", "Hasil rumus Excel: subtotal (SUM), total, surplus/defisit, pembiayaan netto, SILPA, capaian, selisih, perubahan, sisa anggaran, porsi, serta semua angka di Dashboard dan Data Grafik."),
        ("Capaian (%)", "Realisasi dibagi Anggaran. Pada baris Surplus/Defisit tanda negatif mengikuti angka sumber."),
        ("Sisa / (Lebih) Anggaran", "Anggaran dikurangi Realisasi. Untuk belanja, positif berarti anggaran belum terserap. Untuk pendapatan, negatif (dalam kurung) berarti realisasi melampaui target."),
        ("Realisasi tahun lalu", f"Kolom G memakai angka pembanding {Y-1} yang dimuat di LRA {Y}."),
        ("Pembulatan", "LRA 2023 dimuat dalam rupiah penuh tanpa sen, sehingga jumlah rumus dapat berbeda 1 rupiah dari angka total di dokumen sumber. LRA 2024 dan 2025 memuat dua desimal."),
        ("Catatan dokumen 2023", "Pada LRA 2023, anggaran Pencairan Dana Cadangan tertulis 14.940.979.980 sedangkan realisasinya 14.940.879.980. Angka disalin apa adanya sesuai dokumen sumber."),
        ("Cara membaca rumus", "Klik sel hasil, lihat bilah rumus di atas (Excel, Google Sheets, atau WPS). Jika pratinjau HP tidak punya bilah rumus, lihat baris 3 di atas kolom, sheet Rumus & Perhitungan, atau menu Lembar Kerja di dashboard web."),
    ]
    hdr(cat, 4, ["Hal", "Keterangan"])
    for i, (a, b) in enumerate(notes):
        cat.cell(5 + i, 1, a).font = font(True); cat.cell(5 + i, 2, b).font = font()
        for cc in (1, 2): cat.cell(5 + i, cc).border = BOX; cat.cell(5 + i, cc).alignment = Alignment(wrap_text=True, vertical="top")
        cat.row_dimensions[5 + i].height = 32 if len(b) > 100 else 20
    cat.column_dimensions["A"].width = 26; cat.column_dimensions["B"].width = 120

    os.makedirs(OUT, exist_ok=True)
    p = os.path.join(OUT, f"Laporan-APBD-Kota-Kediri-{Y}.xlsx"); wb.save(p); print(p)

if __name__ == "__main__":
    T = tables()
    for Y in (2023, 2024, 2025): build(Y, T)
