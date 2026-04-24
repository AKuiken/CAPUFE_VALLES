# Lectura de CSVs y calculo de indicadores

import io
import re
import pandas as pd
from config import TARIFAS

def clean_int(val):
    try:
        return str(int(float(str(val).strip())))
    except Exception:
        return str(val).strip()


def hora_to_hms(val):
    """DF: entero HHMMSS -> 'HH:MM:SS'"""
    try:
        s = str(int(float(str(val).strip()))).zfill(6)
        return f"{s[:2]}:{s[2:4]}:{s[4:6]}"
    except Exception:
        return str(val).strip()


def hms_to_dec(t):
    """'HH:MM:SS' -> fraccion decimal del dia"""
    try:
        p = str(t).strip().split(":")
        return (int(p[0]) * 3600 + int(p[1]) * 60 + int(float(p[2]))) / 86400
    except Exception:
        return None


def dec_to_hms(d):
    """Fraccion decimal -> 'HH:MM:SS'"""
    try:
        tot = round(abs(float(d)) * 86400)
        return f"{tot // 3600:02d}:{(tot % 3600) // 60:02d}:{tot % 60:02d}"
    except Exception:
        return ""


def _normalize_col(name: str) -> str:
    """Elimina acentos y basura para uso interno (mantiene espacios)."""
    replacements = {
        "a\u0301": "a", "e\u0301": "e", "i\u0301": "i", "o\u0301": "o", "u\u0301": "u",
        "\xe1": "a", "\xe9": "e", "\xed": "i", "\xf3": "o", "\xfa": "u",
        "\xc1": "A", "\xc9": "E", "\xcd": "I", "\xd3": "O", "\xda": "U",
        "\xf1": "n", "\xd1": "N",
        "\ufeff": "",
    }
    s = str(name)
    for old, new in replacements.items():
        s = s.replace(old, new)
    return s.strip()


def _canon(s: str) -> str:
    """Convierte a forma canónica para comparar nombres de columnas."""
    s = _normalize_col(s).lower()
    return re.sub(r"[^a-z0-9]", "", s)


def _find_col(df: pd.DataFrame, wanted: str):
    """Encuentra columna aunque tenga acentos, guiones, mayúsculas, etc."""
    w = _canon(wanted)
    for c in df.columns:
        if _canon(c) == w:
            return c
    return None


def _require_col(df: pd.DataFrame, wanted: str) -> str:
    c = _find_col(df, wanted)
    if not c:
        raise ValueError(f"No se encontró columna '{wanted}'. Columnas disponibles: {list(df.columns)}")
    return c


def read_clr(file_bytes):

    for enc in ("latin-1", "cp1252", "utf-8-sig", "utf-8"):
        try:
            preview = io.BytesIO(file_bytes).read().decode(enc)
            print("=== PRIMERAS 10 LÍNEAS DEL ARCHIVO ===")
            for i, line in enumerate(preview.splitlines()[:10]):
                print(f"  Línea {i}: {line}")
            break
        except Exception:
            continue
    """
    Lee el CSV DF/CLR.
    - 3 filas de titulo antes del encabezado real.
    - Encoding variable.
    - Soporta separador ',' o ';'
    - Primera columna sin nombre = Operador Carretero.
    """

    CANON_TARGETS = {
        "horatransaccion": "Hora Transaccion",
        "fechatransaccion": "Fecha Transaccion",
        "numerotarjeta": "Numero Tarjeta",
        "numerodetransaccion": "Numero de Transaccion",
        "plazadecobro": "Plaza de Cobro",
        "evento": "Evento",
        "carril": "Carril",
        "tramo": "Tramo",
        "clase": "Clase",
        "tipo": "Tipo",
    }

    for enc in ("latin-1", "cp1252", "utf-8-sig", "utf-8"):
        for sep in (",", ";"):
            try:

                skip = next(
    (i for i, line in enumerate(io.BytesIO(file_bytes).read().decode(enc).splitlines()[:10])
     if len(line.split(sep)) > 5),
    3
)
                df = pd.read_csv(
    io.BytesIO(file_bytes),
    encoding=enc,
    sep=sep,
    skiprows=skip,
    dtype=str,
)

                if df.shape[1] <= 1:
                    continue

                # limpia headers
                df.columns = df.columns.astype(str).str.strip().str.lstrip("\ufeff")

                # primera col vacía => Operador Carretero
                if df.columns[0].startswith("Unnamed") or df.columns[0] == "":
                    df.rename(columns={df.columns[0]: "Operador Carretero"}, inplace=True)

                # normaliza acentos en nombres
                df.rename(columns={c: _normalize_col(c) for c in df.columns}, inplace=True)

                # mapea variantes a nombres estándar
                rename_map = {}
                for c in df.columns:
                    k = _canon(c)
                    if k in CANON_TARGETS and CANON_TARGETS[k] not in df.columns:
                        rename_map[c] = CANON_TARGETS[k]
                if rename_map:
                    df.rename(columns=rename_map, inplace=True)

                return df
            except Exception:
                continue

    raise ValueError("No se pudo leer el archivo CLR/DF.")


def read_cci(file_bytes):
    """
    Lee el CSV CCI.
    - Encoding variable.
    - Soporta separador ',' o ';'
    - Intenta skiprows para archivos con encabezados extra.
    """
    best_df = None
    best_cols = 0

    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        for sep in (",", ";"):
            for skip in (0, 1, 2, 3):
                try:
                    df = pd.read_csv(
                        io.BytesIO(file_bytes),
                        encoding=enc,
                        sep=sep,
                        skiprows=skip,
                        dtype=str,
                    )
                    if df.shape[1] <= 1:
                        continue
                    df.columns = df.columns.astype(str).str.strip().str.lstrip("\ufeff")

                    # Preferir la versión con columnas CCI reconocidas
                    canon_cols = set(_canon(c) for c in df.columns)
                    cci_keys = {"uuid", "numerotag", "numerotag", "eventocci"}
                    if cci_keys & canon_cols:
                        return df  # coincidencia exacta, retornar de inmediato

                    # Guardar la mejor opción (más columnas) como fallback
                    if df.shape[1] > best_cols:
                        best_cols = df.shape[1]
                        best_df = df
                except Exception:
                    continue

    if best_df is not None and best_cols >= 5:
        return best_df

    raise ValueError(
        "No se pudo leer el archivo CCI. "
        "Verifique que el archivo correcto esté en el slot CCI "
        "(Datos de Cobro / Dispersión CCI, no un archivo CLR/DF)."
    )


def _clean_clr_numerics(df):
    """Convierte '3421.000000' -> '3421' en columnas numericas del DF."""
    int_cols = ["Carril", "Evento", "Tramo", "Plaza de Cobro",
                "Numero de Transaccion", "Estado Transaccion"]
    for col in int_cols:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda v: clean_int(v) if pd.notna(v) and str(v).strip() not in ("", "nan") else ""
            )
    return df

def build_indicators(clr, cci):
    clr = _clean_clr_numerics(clr.copy())
    cci = cci.copy()

    # -----------------------------
    # CLR
    # -----------------------------
    clr["CONVERSION_HORA"] = clr["Hora Transaccion"].apply(hora_to_hms)
    clr["_dec_clr"] = clr["CONVERSION_HORA"].apply(hms_to_dec)

    clr["CLASE_VAL"] = (
        clr["Clase"].fillna("").str.strip()
        + clr["Tipo"].fillna("").str.strip()
    )

    clr["INDICADOR_TARIFA"] = (
        clr["Tramo"].fillna("").str.strip()
        + clr["CLASE_VAL"]
    )

    clr["IMPORTE_VALUADO"] = clr["INDICADOR_TARIFA"].map(TARIFAS)

    clr["INDICADOR_PISO"] = (
        clr["Numero Tarjeta"].fillna("").str.strip()
        + clr["Evento"].fillna("").str.strip()
        + clr["_dec_clr"].apply(lambda x: f"{x:.6f}" if x is not None else "")
    )

    # -----------------------------
    # CCI
    # -----------------------------
    cci["_dec_cci"] = cci["HORA"].apply(hms_to_dec)

    cci["INDICADOR_CCI"] = (
        cci["NUMERO_TAG"].fillna("").str.strip()
        + cci["EVENTO"].fillna("").str.strip()
        + cci["_dec_cci"].apply(lambda x: f"{x:.6f}" if x is not None else "")
    )

    cci = cci.rename(columns={
        "EVENTO": "EVENTO_CCI",
        "CLASE": "CLASE_CCI"
    })
    
    mg = clr.merge(
    cci,
    left_on="INDICADOR_PISO",
    right_on="INDICADOR_CCI",
    how="left",
    indicator=True   
    )
   # DEBUG (para ver si sí cruza)
    print("CLR:", len(clr))
    print("CCI:", len(cci))
    print("MERGE:", len(mg))
    print("COINCIDENCIAS:", mg["NUMERO_TAG"].notna().sum())

    # -----------------------------
    # VALIDACIONES
    # -----------------------------
    mg["VAL_TARJETA"] = (
        mg["Numero Tarjeta"].fillna("").str.strip()
        == mg["NUMERO_TAG"].fillna("").str.strip()
    )

    mg["VAL_EVENTO"] = (
        mg["Evento"].fillna("").str.strip()
        == mg["EVENTO_CCI"].fillna("").str.strip()
    )

    mg["VAL_TRAMO"] = (
        mg["Tramo"].fillna("").str.strip()
        == mg["CVE_TRAMO"].fillna("").str.strip()
    )

    mg["VAL_CARRIL"] = (
        mg["Carril"].fillna("").str.strip()
        == mg["CVE_CARRIL"].fillna("").str.strip()
    )

    # -----------------------------
    # DIFERENCIAS
    # -----------------------------
    def _val_hora(row):
        try:
            return dec_to_hms(float(row["_dec_clr"]) - float(row["_dec_cci"]))
        except:
            return ""

    def _diferencia(row):
        try:
            v = float(row["IMPORTE_VALUADO"]) if pd.notna(row["IMPORTE_VALUADO"]) else 0
            t = float(row["IMPORTE_TOTAL"]) if pd.notna(row["IMPORTE_TOTAL"]) else 0
            return v - t
        except:
            return None

    mg["VAL_HORA"] = mg.apply(_val_hora, axis=1)
    mg["DIFERENCIA"] = mg.apply(_diferencia, axis=1)

    return mg