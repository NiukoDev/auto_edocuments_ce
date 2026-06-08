APP_NAME = "Auditor de Expedientes"

DATA_SHEET_NAME = "Datos Generales"
DS_SHEET_NAME = "DS"

REQUIRED_DATA_COLUMNS = [
    "Patente",
    "Pedimento",
    "SeccionAduanera",
    "TipoOperacion",
    "ClaveDocumento",
    "FechaPagoReal",
]

DS_COLUMNS = [
    "Patente",
    "Pedimento",
    "SeccionAduanera",
    "LLAVE",
    "TipoOperacion",
    "ClaveDocumento",
    "FechaPagoReal",
    "ESTATUS",
    "MES",
]

MONTHLY_COLUMNS = [
    "PEDIMENTO",
    "CLAVE",
    "PAGO",
    "PD",
    "PS",
    "XMLPD",
    "DODA/PITA/AVC",
    "DETALLE COVE",
    "ACUSE COVE",
    "XML COVE",
    "FACTURA",
    "RNNAS",
    "ESCRITOS",
    "ORIGEN",
    "DOC DE TRANSPORTE",
    "FACTURAS",
    "OTROS DOC",
    "ACUSE",
    "HC",
    "MV",
    "ARCHIVOS DE VALIDACION DE PAGO",
    "Comprobante incrementables",
    "ESTATUS",
    "Comentarios",
]

DOCUMENT_RULES = {
    "PD": {"pattern": "tipodocumento=ped_no=2", "extension": ".pdf"},
    "PS": {"pattern": "tipodocumento=ped_no=1", "extension": ".pdf"},
    "XMLPD": {"pattern": "tipodocumento=xmlc", "extension": ".xml"},
    "DODA/PITA/AVC": {"pattern": "tipodocumento=dpdf", "extension": ".pdf"},
    "DETALLE COVE": {"pattern": "tipodocumento=acvp_no=cove", "extension": ".pdf"},
    "ACUSE COVE": {"pattern": "tipodocumento=acvpa", "extension": ".pdf"},
    "XML COVE": {"pattern": "cove", "extension": ".xml"},
    "FACTURA": {"pattern": "tipodocumento=fac_no=1", "extension": ".pdf"},
    "RNNAS":    {"pattern": "tipodocumento=ed_no=0436", "extension": ".pdf"},
    "ESCRITOS": {"pattern": "tipodocumento=ed_no=0192", "extension": ".pdf"},
    "ORIGEN":          {"pattern": "tipodocumento=ed_no=0433", "extension": ".pdf"},
    "DOC DE TRANSPORTE": {"pattern": "tipodocumento=ed_no=0438", "extension": ".pdf"},
    "FACTURAS":  {"pattern": "tipodocumento=ed_no=0170", "extension": ".pdf"},
    "OTROS DOC": {"pattern": "tipodocumento=ed_no=0172", "extension": ".pdf"},
    "ACUSE":     {"pattern": "tipodocumento=aced_no=",  "extension": ".pdf"},
    "HC":        {"pattern": "hojadecalculo",            "extension": ".pdf"},
    "MV":        {"pattern": "manisfestaciondevalor",    "extension": ".pdf"},
    "Comprobante incrementables": {"pattern": "tipodocumento=cga_no", "extension": ""},
}

DOCUMENT_PATTERNS_BY_PATENT = {
    "3183": {
        "PD": {"doc=desd"},
        "PS": {"doc=simp"},
        "XMLPD": {"doc=cgmx"},
        "DODA/PITA/AVC": {"doc=dpdf"},
        "DETALLE COVE": {"doc=cove"},
        "ACUSE COVE": {"doc=acvpa_arch=acuse"},
        "XML COVE": {"doc=cxml"},
        "FACTURA": {"doc=fac"},
        "ESCRITOS":          {"_doc=edoc_arch=3.18"},
        "ORIGEN":            {"_doc=edoc_arch=origen"},
        "DOC DE TRANSPORTE": {"_doc=edoc_arch=bl"},
        "FACTURAS":          {"_doc=edoc_arch=factura"},
        "ACUSE":             {"doc=aced_arch=acuse_"},
        "HC":                {"doc=hce_arch"},
        "MV":                {"doc=mde"},
        "Comprobante incrementables": {"_doc=cga_arch"},
    },
    "3685": {
        "PD": {"doc=desd"},
        "PS": {"doc=simp"},
        "XMLPD": {"doc=cgmx"},
        "DODA/PITA/AVC": {"doc=dpdf"},
        "DETALLE COVE": {"doc=cove"},
        "ACUSE COVE": {"doc=acvpa_arch=acuse"},
        "XML COVE": {"doc=cxml"},
        "FACTURA": {"doc=fac"},
        "ESCRITOS":          {"_doc=edoc_arch=3.18"},
        "ORIGEN":            {"_doc=edoc_arch=origen"},
        "DOC DE TRANSPORTE": {"_doc=edoc_arch=bl"},
        "FACTURAS":          {"_doc=edoc_arch=factura"},
        "ACUSE":             {"doc=aced_arch=acuse_"},
        "HC":                {"doc=hce_arch"},
        "MV":                {"doc=mde"},
        "Comprobante incrementables": {"_doc=cga_arch"},
    },
    "3908": {
        "ACUSE COVE":        {"acusevalcom_cove"},
        "DOC DE TRANSPORTE": {"bl_"},
        "HC":                {"hojacalculo"},
        "MV":                {"manifestacionvalor_"},
    },
    "1410": {
        "OTROS DOC": {"tipodocumento=ed_no=0428"},
        "HC":        {"tipodocumento=hdc"},
        "MV":        {"tipodocumento=mdv"},
    },
}

DOCUMENT_PREFIX_FALLBACKS = {
    "ACUSE COVE":   {"prefix": "acuse_cove",  "extension": ".pdf"},
    "DETALLE COVE": {"prefix": "cove",         "extension": ".pdf"},
    "XML COVE":     {"prefix": "cove",         "extension": ".xml"},
    "ACUSE":        {"prefix": "acuse_",       "extension": ".pdf"},
}

ACUSE_DEPENDENT_COLUMNS = frozenset({
    "RNNAS", "ESCRITOS", "ORIGEN", "DOC DE TRANSPORTE", "FACTURAS", "OTROS DOC",
})

DOCUMENTS_NA_EXCEPT_PATENTS = {
    "RNNAS":     {"1410"},
    "ORIGEN":    {"1410", "3183", "3685"},
    "FACTURAS":  {"1410", "3183", "3685"},
    "OTROS DOC": {"1410"},
}

DOCUMENTS_NA_FOR_PATENTS = {
    "ESCRITOS": {"3908"},
    "Comprobante incrementables": {"3908"},
}

DOCUMENTS_NA_FOR_CLAVES = {
    "DOC DE TRANSPORTE": {"F4", "F5", "A3", "V1"},
}

DOCUMENTS_NOT_APPLICABLE_BY_COLUMN = {
    "DODA/PITA/AVC": {"Importacion": {"F4", "V1"}},
    "DETALLE COVE": {"Importacion": {"V1"}},
    "ACUSE COVE": {"Importacion": {"V1"}},
    "XML COVE": {"Importacion": {"V1"}},
    "HC": {"Exportacion": {"*"}},
    "MV": {"Exportacion": {"*"}},
    "Comprobante incrementables": {"Importacion": {"V1"}, "Exportacion": {"*"}},
}

PERCENTAGES_SHEET_NAME = "Porcentajes"
PERCENTAGES_COLUMNS = ["MES", "TOTAL", "COMPLETO", "PENDIENTE", "% COMPLETO"]

PAYMENT_VALIDATION_FILES = {
    "1410": [
        {"pattern": "tipodocumento=amm_no=99_arch=e", "extension": "numeric3", "label": "AMM"},
        {"pattern": "tipodocumento=ara_no=1_arch=a",  "extension": "numeric3", "label": "ARA"},
        {"pattern": "tipodocumento=arm_no=1_arch=m",  "extension": "numeric3", "label": "ARM"},
        {"pattern": "tipodocumento=err_no=99_arch=m", "extension": ".err",     "label": "ERR"},
    ],
    "3183": [
        {"pattern": "doc=arca_arch=a", "extension": "numeric3", "label": "ARCA"},
        {"pattern": "doc=arm_arch=m",  "extension": "numeric3", "label": "ARM"},
        {"pattern": "doc=arce_arch=e", "extension": "numeric3", "label": "ARCE"},
        {"pattern": "doc=are_arch=m",  "extension": ".err",     "label": "ARE"},
    ],
    "3685": [
        {"pattern": "doc=arca_arch=a", "extension": "numeric3", "label": "ARCA"},
        {"pattern": "doc=arm_arch=m",  "extension": "numeric3", "label": "ARM"},
        {"pattern": "doc=arce_arch=e", "extension": "numeric3", "label": "ARCE"},
        {"pattern": "doc=are_arch=m",  "extension": ".err",     "label": "ARE"},
    ],
    "3908": [
        {"regex": r"a\d{7}\.\d{3}$", "label": "A#######.###"},
        {"regex": r"e\d{7}\.\d{3}$", "label": "E#######.###"},
        {"regex": r"m\d{7}\.\d{3}$", "label": "M#######.###"},
        {"regex": r"m\d{7}\.err$",   "label": "M#######.ERR"},
    ],
}

MONTHS = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
]
