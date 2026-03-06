# ─────────────────────────────────────────────────────────────────────────────
# CAPUFE – Generación del archivo Excel de conciliación
# ─────────────────────────────────────────────────────────────────────────────

import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from config import ALL_COLS, SECTION_COLORS, COLUMN_WIDTHS, DEFAULT_WIDTH, TARIFAS


def _fill(c): return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=10): return Font(name="Arial", bold=bold, color=color, size=size)
def _aln(wrap=False): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


# ── Hoja CONCILIACION ─────────────────────────────────────────────────────────

def _write_conciliacion(ws, df: pd.DataFrame):
    # Fila 1: etiquetas de sección (fusionadas)
    bounds = {}
    for ci, (_, _, sec, _) in enumerate(ALL_COLS, 1):
        bounds.setdefault(sec, [ci, ci])
        bounds[sec][1] = ci

    for sec, (s, e) in bounds.items():
        c = ws.cell(1, s, sec)
        c.font      = _font(True, SECTION_COLORS[sec]["hfg"], 11)
        c.fill      = _fill(SECTION_COLORS[sec]["hbg"])
        c.alignment = _aln()
        if s != e:
            ws.merge_cells(start_row=1, start_column=s, end_row=1, end_column=e)

    # Fila 2: encabezados
    for ci, (_, label, sec, _) in enumerate(ALL_COLS, 1):
        c = ws.cell(2, ci, label)
        c.font      = _font(True, "000000", 9)
        c.fill      = _fill(SECTION_COLORS[sec]["dbg"])
        c.alignment = _aln(wrap=True)
    ws.row_dimensions[2].height = 34

    # Datos
    for ri, (_, row) in enumerate(df.iterrows(), 3):
        odd = (ri % 2 == 0)
        for ci, (field, _, sec, is_calc) in enumerate(ALL_COLS, 1):
            raw = row.get(field, "")
            val = "" if (raw is None or (not isinstance(raw, bool) and pd.isna(raw))) else raw
            cell = ws.cell(ri, ci)
            cell.alignment = _aln()

            if field == "DIFERENCIA":
                try:
                    n = float(val)
                    cell.value         = n
                    cell.number_format = "#,##0.00"
                    neg = n < 0
                    cell.font = _font(bold=True, color="CC0000" if neg else "006100")
                    cell.fill = _fill("FFCCCC" if neg else "CCFFCC")
                except Exception:
                    cell.value = val; cell.font = _font(bold=True); cell.fill = _fill("FCE4D6")

            elif field in ("VAL_TARJETA", "VAL_EVENTO", "VAL_TRAMO", "VAL_CARRIL"):
                ok = str(val).strip().lower() in ("true", "verdadero", "1")
                cell.value = "VERDADERO" if ok else "FALSO"
                cell.font  = _font(bold=True, color="006100" if ok else "9C0006")
                cell.fill  = _fill("C6EFCE" if ok else "FFC7CE")

            elif field == "IMPORTE_VALUADO":
                try:
                    n = float(val); cell.value = n; cell.number_format = "#,##0"
                except Exception:
                    cell.value = val
                cell.font = _font(bold=True); cell.fill = _fill("FCE4D6")

            elif is_calc:
                cell.value = val
                cell.font  = _font(bold=True)
                cell.fill  = _fill("FFFF99")

            else:
                cell.value = val
                cell.font  = _font()
                cell.fill  = _fill(SECTION_COLORS[sec]["dbg"] if odd else "FFFFFF")

    # Anchos
    for ci, (_, label, _, _) in enumerate(ALL_COLS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COLUMN_WIDTHS.get(label, DEFAULT_WIDTH)

    ws.freeze_panes = "A3"


# ── Hoja TARIFAS ──────────────────────────────────────────────────────────────

def _write_tarifas(ws):
    for ci, txt in enumerate(("CLAVE", "IMPORTE"), 1):
        c = ws.cell(1, ci, txt)
        c.font = _font(True, "FFFFFF", 10); c.fill = _fill("1F4E79"); c.alignment = _aln()

    for i, (k, v) in enumerate(TARIFAS.items(), 2):
        ws.cell(i, 1, k).alignment = _aln(); ws.cell(i, 1).font = _font()
        ws.cell(i, 2, v).alignment = _aln(); ws.cell(i, 2).font = _font()
        ws.cell(i, 2).number_format = "#,##0"
        if i % 2 == 0:
            ws.cell(i, 1).fill = _fill("DEEAF1"); ws.cell(i, 2).fill = _fill("DEEAF1")

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 12


# ── Función pública ───────────────────────────────────────────────────────────

def generate_excel(df: pd.DataFrame) -> io.BytesIO:
    wb = Workbook()

    ws_main = wb.active
    ws_main.title = "CONCILIACION"
    _write_conciliacion(ws_main, df)

    ws_tar = wb.create_sheet("TARIFAS")
    _write_tarifas(ws_tar)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
