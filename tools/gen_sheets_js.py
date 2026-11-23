"""Ekspor sheet Excel (rumus + nilai) ke js/sheets.js untuk tampilan Lembar Kerja di web.
Jalankan setelah build_xlsx.py dan penghitungan ulang (recalc) selesai."""
import os, json
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
DL = os.path.join(HERE, "..", "downloads")
SHEETS = ["Rincian LRA", "Rumus & Perhitungan", "Data Grafik"]
MAXC = {"Data Grafik": 4}                      # kolom rumus (F-H) tidak ditampilkan; rumus terlihat di bilah rumus
FILL = {"22265C": "h", "ECEBF7": "sub", "F6F4EF": "grp"}

def fmt_kind(nf):
    return {'#,##0.00;(#,##0.00);"-"': "r", '0.00%;(0.00%);"-"': "p", "#,##0.00": "m", "0.00%": "q",
            "+0.00%;-0.00%;0.00%": "s", "0.0%": "o"}.get(nf, "")

def cell_json(c, cv):
    if c.value is None: return None
    f = c.value if (c.data_type == "f") else None
    v = cv.value
    st = []
    rgb = (c.fill.start_color.rgb or "")[-6:] if c.fill and c.fill.fill_type == "solid" else ""
    if rgb in FILL: st.append(FILL[rgb])
    fc = ((c.font.color.rgb or "")[-6:] if c.font and c.font.color and isinstance(c.font.color.rgb, str) else "")
    if fc == "0000FF": st.append("in")
    if c.font and c.font.name == "Consolas": st.append("code")
    if c.font and c.font.b: st.append("b")
    if c.font and c.font.i: st.append("i")
    al = c.alignment
    if al and al.horizontal == "center": st.append("ctr")
    if al and al.indent: st.append("ind")
    if al and al.wrap_text: st.append("wrap")
    if isinstance(v, float): v = round(v, 10)
    out = [v, f, fmt_kind(c.number_format), " ".join(st)]
    while out and out[-1] in (None, ""): out.pop()
    return out

def sheet_json(wf, wv, name):
    ws, wsv = wf[name], wv[name]
    mc = MAXC.get(name, ws.max_column)
    rows = []
    for r in range(1, ws.max_row + 1):
        row = [cell_json(ws.cell(r, c), wsv.cell(r, c)) for c in range(1, mc + 1)]
        while row and row[-1] is None: row.pop()
        rows.append(row)
    merges = [[m.min_row, m.min_col, m.max_row, m.max_col] for m in ws.merged_cells.ranges if m.min_col <= mc]
    for m in merges: m[3] = min(m[3], mc)
    widths = [round(ws.column_dimensions[get_column_letter(c)].width or 10, 1) for c in range(1, mc + 1)]
    return {"rows": rows, "merges": merges, "w": widths}

def main():
    out = {}
    for y in (2023, 2024, 2025):
        p = os.path.join(DL, f"Laporan-APBD-Kota-Kediri-{y}.xlsx")
        wf, wv = load_workbook(p), load_workbook(p, data_only=True)
        out[y] = {n: sheet_json(wf, wv, n) for n in SHEETS}
    js = "/* Dibuat otomatis oleh tools/gen_sheets_js.py dari file Excel di downloads/. Jangan diubah manual. */\nwindow.SHEETS = " + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";\n"
    with open(os.path.join(HERE, "..", "js", "sheets.js"), "w", encoding="utf-8") as fh: fh.write(js)
    print("sheets.js", len(js) // 1024, "KB")

if __name__ == "__main__": main()
