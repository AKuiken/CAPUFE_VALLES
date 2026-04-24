# CAPUFE - Configuracion general
TARIFAS = {
    # ── TRAMO 574 ─────────────────────────────────────────────────
    "574T01A": 128, "574T01M": 64,
    "574T02B": 242, "574T02C": 242,
    "574T03B": 242, "574T03C": 242,
    "574T04B": 242, "574T04C": 242,
    "574T05C": 375, "574T06C": 375,
    "574T07C": 457, "574T08C": 457, "574T09C": 457,
    "574EEL":   64, "574EEP":  121,
    # Larga estancia 574
    "574T01L01A": 192, "574T01L02A": 256, "574T01L03A": 320,
    "574T01L04A": 384, "574T01L05A": 448, "574T01L06A": 512,
    "574T01L07A": 576, "574T01L08A": 640, "574T01L09A": 704,
    "574T01L10A": 768, "574T01L11A": 832, "574T01L12A": 896,
    "574T01L13A": 960, "574T01L14A":1024, "574T01L15A":1088,
    # Paso 574
    "574T09P01C":  578, "574T09P02C":  699, "574T09P03C":  820,
    "574T09P04C":  941, "574T09P05C": 1062, "574T09P06C": 1183,
    "574T09P07C": 1304, "574T09P08C": 1425, "574T09P09C": 1546,
    "574T09P10C": 1667, "574T09P11C": 1788, "574T09P12C": 1909,
    "574T09P13C": 2030, "574T09P14C": 2151, "574T09P15C": 2272,
    # ── TRAMO 575 ─────────────────────────────────────────────────
    "575T01A":  52, "575T01M":  26,
    "575T02B": 101, "575T02C": 101,
    "575T03B": 101, "575T03C": 101,
    "575T04B": 101, "575T04C": 101,
    "575T05C": 154, "575T06C": 154,
    "575T07C": 190, "575T08C": 190, "575T09C": 190,
    "575EEL":   26, "575EEP":   51,
    # Larga estancia 575
    "575T01L01A":  78, "575T01L02A": 104, "575T01L03A": 130,
    "575T01L04A": 156, "575T01L05A": 182, "575T01L06A": 208,
    "575T01L07A": 234, "575T01L08A": 260, "575T01L09A": 286,
    "575T01L10A": 312, "575T01L11A": 338, "575T01L12A": 364,
    "575T01L13A": 390, "575T01L14A": 416, "575T01L15A": 442,
    # Paso 575
    "575T09P01C": 241, "575T09P02C": 292, "575T09P03C": 343,
    "575T09P04C": 394, "575T09P05C": 445, "575T09P06C": 496,
    "575T09P07C": 547, "575T09P08C": 598, "575T09P09C": 649,
    "575T09P10C": 700, "575T09P11C": 751, "575T09P12C": 802,
    "575T09P13C": 853, "575T09P14C": 904, "575T09P15C": 955,
    # ── TRAMO 576 ─────────────────────────────────────────────────
    "576T01A":  74, "576T01M":  37,
    "576T02B": 141, "576T02C": 141,
    "576T03B": 141, "576T03C": 141,
    "576T04B": 141, "576T04C": 141,
    "576T05C": 221, "576T06C": 221,
    "576T07C": 267, "576T08C": 267, "576T09C": 267,
    "576EEL":   37, "576EEP":   71,
    # Larga estancia 576
    "576T01L01A": 111, "576T01L02A": 148, "576T01L03A": 185,
    "576T01L04A": 222, "576T01L05A": 259, "576T01L06A": 296,
    "576T01L07A": 333, "576T01L08A": 370, "576T01L09A": 407,
    "576T01L10A": 444, "576T01L11A": 481, "576T01L12A": 518,
    "576T01L13A": 555, "576T01L14A": 592, "576T01L15A": 629,
    # Paso 576
    "576T09P01C":  338, "576T09P02C":  409, "576T09P03C":  480,
    "576T09P04C":  551, "576T09P05C":  622, "576T09P06C":  693,
    "576T09P07C":  764, "576T09P08C":  835, "576T09P09C":  906,
    "576T09P10C":  977, "576T09P11C": 1048, "576T09P12C": 1119,
    "576T09P13C": 1190, "576T09P14C": 1261, "576T09P15C": 1332,
}

DF_COLS = [
    ("Operador Carretero",      "Operador Carretero",       "DF", False),
    ("Fecha Base",              "Fecha Base",               "DF", False),
    ("Plaza de Cobro",          "Plaza de Cobro",           "DF", False),
    ("Numero de Transaccion",   "Num. Transaccion",         "DF", False),
    ("Estado Transaccion",      "Estado Transaccion",       "DF", False),
    ("Descripcion",             "Descripcion",              "DF", False),
    ("Fecha Transaccion",       "Fecha Transaccion",        "DF", False),
    ("Hora Transaccion",        "Hora Transaccion",         "DF", False),
    ("Numero Tarjeta",          "Numero Tarjeta",           "DF", False),
    ("Estado Tarjeta",          "Estado Tarjeta",           "DF", False),
    ("Carril",                  "Carril",                   "DF", False),
    ("Evento",                  "Evento",                   "DF", False),
    ("Clase",                   "Clase",                    "DF", False),
    ("Tipo",                    "Tipo",                     "DF", False),
    ("Tramo",                   "Tramo",                    "DF", False),
    ("Exento",                  "Exento",                   "DF", False),
    ("Exentovig",               "Exentovig",                "DF", False),
    ("Tipo de Usuario",         "Tipo de Usuario",          "DF", False),
    ("Descuento Vigente",       "Descuento Vigente",        "DF", False),
    ("Ind Tipo Respaldo",       "Ind. Tipo Respaldo",       "DF", False),
    ("Forma de pago Capufe",    "Forma de pago Capufe",     "DF", False),
    ("Forma de pago proveedor", "Forma de pago proveedor",  "DF", False),
    ("Descuento TILH",          "Descuento TILH",           "DF", False),
    ("Numero de Cliente",       "Num. Cliente",             "DF", False),
    ("Tipo de Cliente",         "Tipo de Cliente",          "DF", False),
]

CALC_COLS = [
    ("CONVERSION_HORA",  "CONVERSION_HORA",                       "VALUACION", True),
    ("INDICADOR_PISO",   "INDICADOR_ PISO\n(TAG&EVENTO&HORA)",    "VALUACION", True),
    ("CLASE_VAL",        "CLASE",                                 "VALUACION", True),
    ("INDICADOR_TARIFA", "INDICADOR_ TARIFA\n(TRAMO&CLASE&TIPO)", "VALUACION", True),
    ("IMPORTE_VALUADO",  "IMPORTE_VALUADO",                       "VALUACION", True),
    ("VAL_TARJETA",      "VAL_TARJETA",                           "VALUACION", True),
    ("VAL_EVENTO",       "VAL_EVENTO",                            "VALUACION", True),
    ("VAL_TRAMO",        "VAL TRAMO",                             "VALUACION", True),
    ("VAL_CARRIL",       "VAL_CARRIL",                            "VALUACION", True),
    ("VAL_HORA",         "VAL_HORA",                              "VALUACION", True),
    ("DIFERENCIA",       "DIFERENCIA",                            "VALUACION", True),
]

CCI_COLS = [
    ("UUID",           "UUID",                              "CCI", False),
    ("NUMERO_TAG",     "NUMERO_TAG",                        "CCI", False),
    ("FECHA",          "FECHA",                             "CCI", False),
    ("HORA",           "HORA",                              "CCI", False),
    ("CVE_CARRIL",     "CVE_CARRIL",                        "CCI", False),
    ("SENTIDO",        "SENTIDO",                           "CCI", False),
    ("EVENTO_CCI",     "EVENTO",                            "CCI", False),
    ("CVE_PLAZA",      "CVE_PLAZA",                         "CCI", False),
    ("CLASE_CCI",      "CLASE (CCI)",                       "CCI", False),
    ("CVE_TRAMO",      "CVE_TRAMO",                         "CCI", False),
    ("ID_OPERADOR",    "ID_OPERADOR",                       "CCI", False),
    ("FECHA_DEPOSITO", "FECHA_DEPOSITO",                    "CCI", False),
    ("HORA_DEPOSITO",  "HORA_DEPOSITO",                     "CCI", False),
    ("ID_ENVIO_OTI",   "ID_ENVIO_OTI",                      "CCI", False),
    ("IMPORTE_TOTAL",  "IMPORTE_TOTAL",                     "CCI", False),
    ("TIPO_TARJETA",   "TIPO_TARJETA",                      "CCI", False),
    ("COMISION",       "COMISION",                          "CCI", False),
    ("INDICADOR_CCI",  "INDICADOR_ CCI\n(TAG&EVENTO&HORA)", "CCI", True),
]

RESULT_COLS = [
    ("ESCENARIO",   "ESCENARIO",   "RESULTADO", False),
    ("DF_DIA",      "DF_DIA",      "RESULTADO", False),
    ("CCI_DIA",     "CCI_DIA",     "RESULTADO", False),
    ("DIF_IMPORTE", "DIF_IMPORTE", "RESULTADO", False),
]
ALL_COLS = RESULT_COLS + DF_COLS + CALC_COLS + CCI_COLS

SECTION_COLORS = {
    "RESULTADO": {"hbg": "0B3D2E", "hfg": "FFFFFF", "dbg": "E8F5E9"},
    "DF":        {"hbg": "1F4E79", "hfg": "FFFFFF", "dbg": "DEEAF1"},
    "VALUACION": {"hbg": "7B3F00", "hfg": "FFFFFF", "dbg": "FCE4D6"},
    "CCI":       {"hbg": "1E4620", "hfg": "FFFFFF", "dbg": "E2EFDA"},
}

COLUMN_WIDTHS = {
    "Descripcion":                            36,
    "INDICADOR_ PISO\n(TAG&EVENTO&HORA)":    45,
    "INDICADOR_ TARIFA\n(TRAMO&CLASE&TIPO)": 28,
    "INDICADOR_ CCI\n(TAG&EVENTO&HORA)":     45,
    "UUID":                                   36,
    "ID_ENVIO_OTI":                           20,
    "Numero Tarjeta":                         16,
    "NUMERO_TAG":                             16,
    "ESCENARIO": 55,
    "DF_DIA": 12,
    "CCI_DIA": 12,
    "DIF_IMPORTE": 14,
}
DEFAULT_WIDTH = 13


