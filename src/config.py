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
