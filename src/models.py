from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class DataRecord:
    patente: str
    pedimento: str
    seccion_aduanera: str
    llave: str
    tipo_operacion: Any
    clave_documento: str
    fecha_pago_real: Any
    estatus: str
    mes: str


@dataclass(frozen=True)
class AnalysisResult:
    data_path: Path
    sheet_name: str
    total_records: int
    duplicate_keys: list[str]
    records: list[DataRecord]


@dataclass(frozen=True)
class GenerationResult:
    output_path: Path
    backup_path: Path | None
    records_written: int
    month_replaced: bool
    generated_at: datetime
    pd_found: int = 0
    pd_missing: int = 0
    ps_found: int = 0
    ps_missing: int = 0
    xmlpd_found: int = 0
    xmlpd_missing: int = 0
