# CAPUFE - Configuracion general
TARIFAS = {
    "574T01A":122,"574T01M":61,"574T02B":221,"574T02C":221,
    "574T03B":221,"574T03C":221,"574T04B":221,"574T04C":221,
    "574T05C":342,"574T06C":342,"574T07C":417,"574T08C":417,
    "574T09C":417,"574T01L01A":183,"574T01L02A":244,"574T01L03A":305,
    "574EEL":61,"574EEP":111,
    "575T01A":51,"575T01M":26,"575T02B":92,"575T02C":92,
    "575T03B":92,"575T03C":92,"575T04B":92,"575T04C":92,
    "575T05C":140,"575T06C":140,"575T07C":173,"575T08C":173,
    "575T09C":173,"575T01L01A":76,"575T01L02A":101,"575T01L03A":126,
    "575EEL":25,"575EEP":46,
    "576T01A":71,"576T01M":35,"576T02B":129,"576T02C":129,
    "576T03B":129,"576T03C":129,"576T04B":129,"576T04C":129,
    "576T05C":202,"576T06C":202,"576T07C":244,"576T08C":244,
    "576T09C":244,"576T01L01A":107,"576T01L02A":143,"576T01L03A":179,
    "576EEL":36,"576EEP":55,
}

# Cada tupla: (campo_en_dataframe, etiqueta_excel, seccion, es_calculado)
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


