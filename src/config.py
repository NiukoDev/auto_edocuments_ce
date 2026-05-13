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

DOCUMENT_PATTERNS = {
    "PD": "tipodocumento=ped_no=2",
    "PS": "tipodocumento=ped_no=1",
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
