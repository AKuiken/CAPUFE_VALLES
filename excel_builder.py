# ─────────────────────────────────────────────────────────────────────────────
# Generación del archivo Excel 
# ─────────────────────────────────────────────────────────────────────────────

import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

from config import ALL_COLS, SECTION_COLORS, COLUMN_WIDTHS, DEFAULT_WIDTH, TARIFAS


def _fill(c): return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=10): return Font(name="Arial", bold=bold, color=color, size=size)
def _aln(wrap=False): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


# ── Hoja CONCILIACION ─────────────────────────────────────────────────────────

def _write_conciliacion(ws, df: pd.DataFrame):
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
            
            if field in df.columns:
                 val = row[field]
            else:
                     val = ""

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
                if field == "Hora Transaccion":
                    try:
                        s = str(val).strip().zfill(6)
                        cell.value = f"{s[:2]}:{s[2:4]}:{s[4:6]}" if s.isdigit() else str(val)
                    except Exception:
                        cell.value = val
                else:
                    cell.value = val
                cell.font  = _font()
                cell.fill  = _fill(SECTION_COLORS[sec]["dbg"] if odd else "FFFFFF")

            # Color por escenario
            if field == "ESCENARIO":
                escenario = str(val).strip()
                if "sin diferencia" in escenario and "coincidentes" in escenario:
                    cell.font = _font(bold=True, color="3B6FCC")
                    cell.fill = _fill("FFFFFF")
                elif "con diferencia" in escenario:
                    cell.font = _font(bold=True, color="FF0000")
                    cell.fill = _fill("FFFFFF")
                elif "día anterior" in escenario or "día posterior" in escenario:
                    cell.font = _font(bold=True, color="1F4E79")
                    cell.fill = _fill("DEEAF1")
                elif "Sobrantes de DF y no en CCI" in escenario:
                    cell.font = _font(bold=True, color="1F4E79")
                    cell.fill = _fill("FFFFFF")
                elif escenario == "Sobrantes de DF y no en CCI, ni en dia anterior o posterior":
                    cell.font = _font(bold=True, color="FF0000")
                    cell.fill = _fill("FFFFFF")

    for ci, (_, label, _, _) in enumerate(ALL_COLS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COLUMN_WIDTHS.get(label, DEFAULT_WIDTH)

    ws.freeze_panes = "A3"


# ── Hoja TARIFAS ──────────────────────────────────────────────────────────────

def _write_tarifas(ws):
    """
    Genera la hoja TARIFAS como tabla cruzada:
      Col A = CLASE, Col B = TRAMO 574, Col C = TRAMO 575, Col D = TRAMO 576
    Con secciones: Básicas, Larga Estancia (L01-L15), Paso Posterior (P01-P15).
    """
    from openpyxl.styles import Border, Side
    from openpyxl.utils import get_column_letter

    TRAMOS = ["574", "575", "576"]

    # Colores por tramo 
    TRAMO_HDR_BG  = {"574": "C55A11",  "575": "375623",  "576": "1F4E79"}
    TRAMO_HDR_FG  = {"574": "FFFFFF",  "575": "FFFFFF",  "576": "FFFFFF"}
    TRAMO_DATA_BG = {"574": "FCE4D6",  "575": "E2EFDA",  "576": "DEEAF1"}

    # Secciones
    SECTIONS = [
        ("Tarifas Básicas", [
            "T01A","T01M",
            "T02B","T02C",
            "T03B","T03C",
            "T04B","T04C",
            "T05C","T06C",
            "T07C","T08C","T09C",
            "EEL","EEP",
        ]),
        ("Larga Estancia (T01L)", [f"T01L{i:02d}A" for i in range(1, 16)]),
        ("Paso Posterior (T09P)", [f"T09P{i:02d}C" for i in range(1, 16)]),
    ]

    thin = Side(style="thin", color="BFBFBF")
    thick = Side(style="medium", color="595959")
    def border(top=None,bottom=None,left=None,right=None):
        return Border(top=top,bottom=bottom,left=left,right=right)

    def _set(ws, row, col, val="", bold=False, fg="000000", bg=None,
             size=9, wrap=False, num_fmt=None, b_top=None, b_bot=None, b_left=None, b_right=None):
        c = ws.cell(row, col, val)
        c.font = Font(name="Arial", bold=bold, color=fg, size=size)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        if bg:
            c.fill = PatternFill("solid", fgColor=bg)
        if num_fmt:
            c.number_format = num_fmt
        c.border = Border(
            top=b_top or thin, bottom=b_bot or thin,
            left=b_left or thin, right=b_right or thin
        )
        return c

    ws.merge_cells("A1:D1")
    c = ws.cell(1, 1, "TARIFAS VIGENTES — PC 0231 · A PARTIR DEL 13/04/2026")
    c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    c.fill      = PatternFill("solid", fgColor="0B3D2E")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    _set(ws, 2, 1, "CLASE", bold=True, fg="FFFFFF", bg="2E4057", size=10,
         b_left=thick, b_top=thick, b_bot=thick)
    for ci, tramo in enumerate(TRAMOS, 2):
        _set(ws, 2, ci, f"TRAMO {tramo}", bold=True,
             fg=TRAMO_HDR_FG[tramo], bg=TRAMO_HDR_BG[tramo], size=10,
             b_top=thick, b_bot=thick,
             b_right=thick if ci == 4 else thin)
    ws.row_dimensions[2].height = 22

    row = 3
    for sec_title, clases in SECTIONS:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        c = ws.cell(row, 1, sec_title)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=9)
        c.fill      = PatternFill("solid", fgColor="595959")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.border    = Border(top=thick, bottom=thick, left=thick, right=thick)
        ws.row_dimensions[row].height = 16
        row += 1

        for idx, clase in enumerate(clases):
            is_even = idx % 2 == 0
            row_bg = "F5F5F5" if is_even else "FFFFFF"

            _set(ws, row, 1, clase, bold=False, fg="1a1a1a", bg=row_bg,
                 b_left=thick, b_right=thin)

            for ci, tramo in enumerate(TRAMOS, 2):
                key = f"{tramo}{clase}"
                val = TARIFAS.get(key, "—")
                is_num = isinstance(val, (int, float))
                _set(ws, row, ci,
                     val if is_num else val,
                     bold=is_num, fg="000000" if is_num else "999999",
                     bg=TRAMO_DATA_BG[tramo] if is_num else row_bg,
                     num_fmt="#,##0" if is_num else None,
                     b_right=thick if ci == 4 else thin)
            row += 1

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 13

    ws.freeze_panes = "B3"

# ──  ──────────────────────────────────────────────────────────────


def _build_scenarios(df_hoy: pd.DataFrame, merged_ant: pd.DataFrame | None, merged_pos: pd.DataFrame | None):
    df = df_hoy.copy()
# DEBUG temporal
    print("=== COLUMNAS EN DF_HOY ===", list(df.columns))
    print("=== FILAS ===", len(df))
    print("=== _merge valores ===", df["_merge"].value_counts() if "_merge" in df.columns else "NO EXISTE")

    df["DF_DIA"] = "HOY"
    df["CCI_DIA"] = "HOY"
    df["DIF_IMPORTE"] = False
    df["ESCENARIO"] = "SIN_CLASIFICAR"

    if "DIFERENCIA" in df.columns:
        df["DIF_IMPORTE"] = df["DIFERENCIA"].astype(str).str.replace(",", "", regex=False)
        df["DIF_IMPORTE"] = df["DIF_IMPORTE"].astype(float).abs().round(2) > 0

    if "_merge" in df.columns:
        mask_both = df["_merge"] == "both"
        df.loc[mask_both & (~df["DIF_IMPORTE"]), "ESCENARIO"] = "Transacciones coincidentes, sin diferencia de importe"
        df.loc[mask_both & ( df["DIF_IMPORTE"]), "ESCENARIO"] = "Transacciones coincidentes, con diferencia de importe"
        df.loc[df["_merge"] == "left_only", "ESCENARIO"] = "Sobrante de DF, localizada en CCI de día anterior"
        df.loc[df["_merge"] == "right_only", "ESCENARIO"] = "Sobrantes de CCI y no en DF"
    # ── Buscar sobrantes en día anterior y posterior ──
    mask_sobrante = df["ESCENARIO"] == "Sobrantes de DF y no en CCI"

    if merged_ant is not None and "INDICADOR_CCI" in merged_ant.columns:
        nums_ant = set(merged_ant["INDICADOR_CCI"].dropna().str.strip())
        mask_en_ant = mask_sobrante & df["INDICADOR_PISO"].isin(nums_ant)
        df.loc[mask_en_ant, "ESCENARIO"] = "Sobrante de DF, localizada en CCI de día anterior"
        df.loc[mask_en_ant, "CCI_DIA"] = "ANTERIOR"

    if merged_pos is not None and "INDICADOR_CCI" in merged_pos.columns:
        nums_pos = set(merged_pos["INDICADOR_CCI"].dropna().str.strip())
        mask_sobrante_aun = df["ESCENARIO"] == "Sobrantes de DF y no en CCI"
        mask_en_pos = mask_sobrante_aun & df["INDICADOR_PISO"].isin(nums_pos)
        df.loc[mask_en_ant, "ESCENARIO"] = "Sobrante de DF, localizada en CCI de día anterior"
        df.loc[mask_en_pos, "CCI_DIA"] = "POSTERIOR"

    ajustes = df[df["DIF_IMPORTE"] == True].copy()

    orden = {
        "Transacciones coincidentes, sin diferencia de importe": 0,
        "Transacciones coincidentes, con diferencia de importe": 1,
        "Sobrantes de DF, localizada en CCI de día anterior, sin diferencia de importe": 2,
        "Sobrantes de DF, localizada en CCI de día posterior, sin diferencia de importe": 3,
        "Sobrantes de DF y no en CCI": 4,
        "Sobrantes de CCI y no en DF": 5,
        "SIN_CLASIFICAR": 6,
    }
    df["_orden"] = df["ESCENARIO"].map(orden).fillna(99)
    df = df.sort_values("_orden").drop(columns=["_orden"]).reset_index(drop=True)

    return df, ajustes


# ── Función pública
def generate_excel(df_hoy: pd.DataFrame, merged_ant=None, merged_pos=None) -> io.BytesIO:
    wb = Workbook()

    conciliacion, ajustes = _build_scenarios(df_hoy, merged_ant, merged_pos)

    ws_main = wb.active
    ws_main.title = "CONCILIACION"
    _write_conciliacion(ws_main, conciliacion)

    ws_aj = wb.create_sheet("AJUSTES")
    _write_conciliacion(ws_aj, ajustes)

    ws_tar = wb.create_sheet("TARIFAS")
    _write_tarifas(ws_tar)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf