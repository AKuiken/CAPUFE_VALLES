# Lectura de CSVs y calculo de indicadores
import io
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


def _normalize_col(name):
    """Elimina acentos para uso interno."""
    replacements = {
        "a\u0301":"a","e\u0301":"e","i\u0301":"i","o\u0301":"o","u\u0301":"u",
        "\xe1":"a","\xe9":"e","\xed":"i","\xf3":"o","\xfa":"u",
        "\xc1":"A","\xc9":"E","\xcd":"I","\xd3":"O","\xda":"U",
        "\xf1":"n","\xd1":"N",".":"",
    }
    for old, new in replacements.items():
        name = name.replace(old, new)
    return name.strip()


def read_clr(file_bytes):
    """
    Lee el CSV DF/CLR.
    - 3 filas de titulo antes del encabezado real.
    - Encoding latin-1.
    - Primera columna sin nombre = Operador Carretero.
    """
    for enc in ("latin-1", "cp1252", "utf-8-sig", "utf-8"):
        try:
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                encoding=enc,
                sep=",",
                skiprows=3,
                dtype=str,
            )
            df.columns = df.columns.str.strip().str.lstrip("\ufeff")
            if df.columns[0].startswith("Unnamed") or df.columns[0] == "":
                df.rename(columns={df.columns[0]: "Operador Carretero"}, inplace=True)
            df.rename(columns={c: _normalize_col(c) for c in df.columns}, inplace=True)
            return df
        except Exception:
            continue
    raise ValueError("No se pudo leer el archivo CLR/DF.")


def read_cci(file_bytes):
    """
    Lee el CSV CCI.
    - Sin filas de titulo.
    - Encoding utf-8-sig.
    """
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                encoding=enc,
                sep=",",
                dtype=str,
            )
            df.columns = df.columns.str.strip().str.lstrip("\ufeff")
            return df
        except Exception:
            continue
    raise ValueError("No se pudo leer el archivo CCI.")


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
    """
    Calcula todos los campos derivados y hace el merge CLR <- CCI.
    Devuelve un DataFrame con todas las columnas de salida.
    """
    clr = _clean_clr_numerics(clr.copy())
    cci = cci.copy()

    # Indicadores CLR
    clr["CONVERSION_HORA"] = clr["Hora Transaccion"].apply(hora_to_hms)
    clr["_dec_clr"]        = clr["CONVERSION_HORA"].apply(hms_to_dec)

    clr["CLASE_VAL"]        = (clr["Clase"].fillna("").str.strip()
                               + clr["Tipo"].fillna("").str.strip())
    clr["INDICADOR_TARIFA"] = (clr["Tramo"].fillna("").str.strip()
                               + clr["CLASE_VAL"])
    clr["IMPORTE_VALUADO"]  = clr["INDICADOR_TARIFA"].map(TARIFAS)

    clr["INDICADOR_PISO"] = (
        clr["Numero Tarjeta"].fillna("").str.strip()
        + clr["Evento"].fillna("").str.strip()
        + clr["_dec_clr"].apply(lambda x: str(x) if x is not None else "")
    )

    # Indicadores CCI
    cci["_dec_cci"] = cci["HORA"].apply(hms_to_dec)
    cci["INDICADOR_CCI"] = (
        cci["NUMERO_TAG"].fillna("").str.strip()
        + cci["EVENTO"].fillna("").str.strip()
        + cci["_dec_cci"].apply(lambda x: str(x) if x is not None else "")
    )
    cci = cci.rename(columns={"EVENTO": "EVENTO_CCI", "CLASE": "CLASE_CCI"})

    # Merge
    mg = clr.merge(cci, left_on="INDICADOR_PISO", right_on="INDICADOR_CCI", how="left")

    # Validaciones
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

    def _val_hora(row):
        try:
            return dec_to_hms(float(row["_dec_clr"]) - float(row["_dec_cci"]))
        except Exception:
            return ""

    def _diferencia(row):
        try:
            v = float(row["IMPORTE_VALUADO"]) if pd.notna(row["IMPORTE_VALUADO"]) else 0.0
            t = float(row["IMPORTE_TOTAL"])   if pd.notna(row["IMPORTE_TOTAL"])   else 0.0
            return v - t
        except Exception:
            return None

    mg["VAL_HORA"]   = mg.apply(_val_hora,   axis=1)
    mg["DIFERENCIA"] = mg.apply(_diferencia, axis=1)

    return mg
