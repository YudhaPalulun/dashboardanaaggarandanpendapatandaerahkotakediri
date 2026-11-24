"""Ekspor angka dari file Excel (nilai hasil rumus) ke js/data.js untuk dashboard web."""
import os, json, sys
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from openpyxl import load_workbook
from data import ROWS, ROW
out = {"years": [2023, 2024, 2025], "rows": [[no, ur, lv, k] for no, ur, lv, k, rule in ROWS], "v": {}}
for y in out["years"]:
    ws = load_workbook(os.path.join(HERE, "..", "downloads", f"Laporan-Pembiayaan-Kota-Kediri-{y}.xlsx"), data_only=True)["Rincian LRA"]
    out["v"][y] = {no: [round(ws.cell(ROW[no], c).value or 0, 2) for c in (4, 5, 7)] for no, *_ in ROWS if no != 42}
js = "/* Dibuat otomatis oleh tools/gen_data_js.py dari file Excel. Jangan diubah manual. */\nwindow.DATA = " + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";\n"
open(os.path.join(HERE, "..", "js", "data.js"), "w", encoding="utf-8").write(js); print("data.js", len(js))
