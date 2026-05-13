from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QComboBox, QCheckBox, QTextEdit, QFileDialog, QMessageBox,
    QVBoxLayout, QHBoxLayout, QFrame, QSizePolicy,
)

from config import APP_NAME, DATA_SHEET_NAME, MONTHS
from excel_service import (
    ExcelValidationError, analyze_data_file, generate_result_file,
    inspect_output_file,
)
from models import AnalysisResult

# ── Paleta ────────────────────────────────────────────────────────────────────
BG     = "#0b1120"
CARD   = "#111827"
PANEL  = "#172033"
FIELD  = "#020617"
TEXT   = "#e5f0ff"
MUTED  = "#94a3b8"
FG_LBL = "#cbd5e1"
CYAN   = "#67e8f9"
BLUE   = "#2563eb"
ACCENT = "#0891b2"
NEUTRAL= "#1f2937"
BORDER = "#334155"

STYLESHEET = f"""
/* ── Base ── */
QMainWindow {{ background: {BG}; }}
QWidget     {{ background: transparent; color: {TEXT}; font-family: Helvetica, Arial, sans-serif; }}

/* ── Cards ── */
QFrame#card  {{ background: {CARD}; border-radius: 8px; }}
QFrame#panel {{ background: {PANEL}; border-radius: 5px; }}
QFrame#sep   {{ background: {BORDER}; }}

/* ── Labels ── */
QLabel {{ background: transparent; color: {FG_LBL}; }}

/* ── ComboBox ── */
QComboBox {{
    background: {FIELD}; color: {TEXT};
    border: 1px solid {BORDER}; border-radius: 4px;
    padding: 5px 10px; font-size: 11px; font-weight: bold;
}}
QComboBox:hover {{ border-color: {CYAN}; }}
QComboBox::drop-down {{ border: none; width: 24px; }}
QComboBox::down-arrow {{
    border-left: 4px solid transparent; border-right: 4px solid transparent;
    border-top: 5px solid {MUTED}; width: 0; height: 0; margin-right: 8px;
}}
QComboBox QAbstractItemView {{
    background: {FIELD}; color: {TEXT};
    selection-background-color: {NEUTRAL};
    border: 1px solid {BORDER}; outline: none; padding: 2px;
}}

/* ── Buttons ── */
QPushButton {{
    border: none; border-radius: 4px;
    padding: 5px 14px; font-weight: bold; font-size: 11px;
}}
QPushButton#neutral  {{ background: {NEUTRAL}; color: {TEXT}; }}
QPushButton#neutral:hover  {{ background: #374151; }}
QPushButton#blue     {{ background: {BLUE}; color: white; }}
QPushButton#blue:hover     {{ background: #1d4ed8; }}
QPushButton#accent   {{ background: {ACCENT}; color: white; }}
QPushButton#accent:hover   {{ background: #0e7490; }}

/* ── Checkbox ── */
QCheckBox {{ color: {FG_LBL}; font-size: 11px; spacing: 8px; }}
QCheckBox::indicator {{
    width: 15px; height: 15px;
    background: {FIELD}; border: 1px solid {BORDER}; border-radius: 3px;
}}
QCheckBox::indicator:checked {{ background: {ACCENT}; border-color: {ACCENT}; }}

/* ── Log text ── */
QTextEdit {{
    background: {FIELD}; color: #dbeafe;
    border: none;
    font-family: Menlo, Monaco, "Courier New", monospace;
    font-size: 10px; padding: 6px;
}}
QScrollBar:vertical {{
    background: {CARD}; width: 7px; border-radius: 3px;
}}
QScrollBar::handle:vertical {{
    background: {BORDER}; border-radius: 3px; min-height: 20px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
"""

# ── Path label estilos ────────────────────────────────────────────────────────
_PATH_DEFAULT = f"background:{FIELD}; color:{MUTED}; font-size:11px; padding:5px 8px; border:1px solid {BORDER}; border-radius:4px;"
_PATH_ACTIVE  = f"background:{FIELD}; color:{TEXT};  font-size:11px; padding:5px 8px; border:1px solid {ACCENT}; border-radius:4px;"


class AuditorApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(980, 700)
        self.setMinimumSize(860, 600)
        self.setStyleSheet(STYLESHEET)

        # Estado
        self._data_path    = ""
        self._output_path  = ""
        self._folder_path  = ""
        self.analysis: AnalysisResult | None = None
        self.log_lines: list[str] = []

        # Widgets dinámicos (se asignan en _build_*)
        self.mode_combo:     QComboBox  | None = None
        self.month_combo:    QComboBox  | None = None
        self.year_combo:     QComboBox  | None = None
        self.overwrite_chk:  QCheckBox  | None = None
        self.log_area:       QTextEdit  | None = None
        self.output_section: QLabel     | None = None
        self.output_btn:     QPushButton| None = None

        self.path_lbl: dict[str, QLabel] = {}

        self._build_ui()
        self._refresh_output_label()
        self._log("Listo. Selecciona el DATA mensual, resultado y carpeta de expedientes.")

    # ── Construcción ─────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("central")
        central.setStyleSheet(f"QWidget#central {{ background: {BG}; }}")
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._header())
        root.addWidget(self._hsep())

        body = QWidget()
        body.setStyleSheet(f"background: {BG};")
        bl = QHBoxLayout(body)
        bl.setContentsMargins(20, 16, 20, 16)
        bl.setSpacing(12)
        bl.addWidget(self._left_panel(), 3)
        bl.addWidget(self._right_panel(), 2)
        root.addWidget(body, 1)

        root.addWidget(self._hsep())
        root.addWidget(self._action_row())

    def _header(self) -> QFrame:
        h = QFrame()
        h.setStyleSheet(f"background: {BLUE};")
        hl = QVBoxLayout(h)
        hl.setContentsMargins(22, 14, 22, 14)
        hl.setSpacing(3)
        t = QLabel(APP_NAME)
        t.setStyleSheet("color: white; font-size: 20px; font-weight: bold;")
        s = QLabel("Genera DS acumulada y hojas mensuales estrictas desde el DATA oficial.")
        s.setStyleSheet("color: #bfdbfe; font-size: 11px;")
        hl.addWidget(t)
        hl.addWidget(s)
        return h

    def _hsep(self) -> QFrame:
        f = QFrame()
        f.setObjectName("sep")
        f.setFixedHeight(1)
        return f

    def _left_panel(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(18, 16, 18, 16)
        cl.setSpacing(0)

        # Modo
        cl.addWidget(self._section("Modo de operación"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Crear archivo nuevo", "Actualizar archivo existente"])
        self.mode_combo.currentTextChanged.connect(self._refresh_output_label)
        cl.addWidget(self.mode_combo)
        cl.addSpacing(14)

        # DATA
        cl.addWidget(self._section("Archivo DATA mensual"))
        lbl_data, row_data = self._path_row("data", "Sin DATA seleccionado",
                                             "Seleccionar DATA", self._select_data)
        cl.addWidget(row_data)
        cl.addSpacing(14)

        # Output
        self.output_section = self._section("Guardar resultado como")
        cl.addWidget(self.output_section)
        lbl_out, row_out = self._path_row("output", "Sin resultado seleccionado",
                                          "Elegir ruta", self._select_output,
                                          capture_btn="output_btn")
        cl.addWidget(row_out)
        cl.addSpacing(14)

        # Carpeta
        cl.addWidget(self._section("Carpeta de expedientes"))
        lbl_folder, row_folder = self._path_row("folder", "Sin carpeta seleccionada",
                                                 "Seleccionar carpeta", self._select_folder)
        cl.addWidget(row_folder)
        cl.addSpacing(14)

        # Mes / Año
        dr = QWidget()
        dr.setStyleSheet(f"background: {CARD};")
        dl = QHBoxLayout(dr)
        dl.setContentsMargins(0, 0, 0, 0)
        dl.setSpacing(10)

        mw = QWidget()
        mw.setStyleSheet(f"background: {CARD};")
        ml = QVBoxLayout(mw)
        ml.setContentsMargins(0, 0, 0, 0)
        ml.setSpacing(4)
        ml.addWidget(self._section("Mes"))
        self.month_combo = QComboBox()
        self.month_combo.addItems(MONTHS)
        self.month_combo.setCurrentIndex(datetime.now().month - 1)
        ml.addWidget(self.month_combo)

        yw = QWidget()
        yw.setStyleSheet(f"background: {CARD};")
        yl = QVBoxLayout(yw)
        yl.setContentsMargins(0, 0, 0, 0)
        yl.setSpacing(4)
        yl.addWidget(self._section("Año"))
        self.year_combo = QComboBox()
        self.year_combo.addItems([str(datetime.now().year + i) for i in range(-2, 4)])
        self.year_combo.setCurrentText(str(datetime.now().year))
        yl.addWidget(self.year_combo)

        dl.addWidget(mw, 3)
        dl.addWidget(yw, 2)
        cl.addWidget(dr)
        cl.addSpacing(14)

        # Info box
        info = QFrame()
        info.setObjectName("panel")
        il = QVBoxLayout(info)
        il.setContentsMargins(12, 10, 12, 10)
        il.setSpacing(4)
        t1 = QLabel(f'El DATA debe tener una hoja llamada exactamente "{DATA_SHEET_NAME}".')
        t1.setStyleSheet(f"color: {CYAN}; font-weight: bold; font-size: 11px;")
        t1.setWordWrap(True)
        t2 = QLabel("Columnas: Patente, Pedimento, SeccionAduanera, TipoOperacion, ClaveDocumento, FechaPagoReal.")
        t2.setStyleSheet(f"color: {FG_LBL}; font-size: 10px;")
        t2.setWordWrap(True)
        il.addWidget(t1)
        il.addWidget(t2)
        cl.addWidget(info)
        cl.addSpacing(12)

        # Checkbox
        self.overwrite_chk = QCheckBox("Permitir reescribir el mes con confirmación")
        cl.addWidget(self.overwrite_chk)
        cl.addStretch()

        return card

    def _right_panel(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(8)

        st = QLabel("Estado")
        st.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: bold;")
        cl.addWidget(st)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        cl.addWidget(self.log_area, 1)

        return card

    def _action_row(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet(f"background: {BG};")
        al = QHBoxLayout(w)
        al.setContentsMargins(20, 10, 20, 12)
        al.setSpacing(10)

        btn_a = QPushButton("Analizar")
        btn_a.setObjectName("blue")
        btn_a.setMinimumHeight(38)
        btn_a.setCursor(Qt.PointingHandCursor)
        btn_a.clicked.connect(self._analyze)

        btn_g = QPushButton("Generar Auditoría")
        btn_g.setObjectName("accent")
        btn_g.setMinimumHeight(38)
        btn_g.setCursor(Qt.PointingHandCursor)
        btn_g.clicked.connect(self._generate)

        al.addWidget(btn_a, 1)
        al.addWidget(btn_g, 1)
        return w

    # ── Widget helpers ────────────────────────────────────────────────────────

    def _section(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color: {FG_LBL}; font-weight: bold; font-size: 11px; margin-bottom: 3px;")
        return lbl

    def _path_row(self, key: str, placeholder: str, btn_text: str, cmd,
                  capture_btn: str | None = None) -> tuple[QLabel, QWidget]:
        row = QWidget()
        row.setStyleSheet(f"background: {CARD};")
        rl = QHBoxLayout(row)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(8)

        lbl = QLabel(placeholder)
        lbl.setStyleSheet(_PATH_DEFAULT)
        lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.path_lbl[key] = lbl

        btn = QPushButton(btn_text)
        btn.setObjectName("neutral")
        btn.setMinimumHeight(32)
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(cmd)

        if capture_btn:
            setattr(self, capture_btn, btn)

        rl.addWidget(lbl, 1)
        rl.addWidget(btn)
        return lbl, row

    # ── Lógica de rutas ───────────────────────────────────────────────────────

    def _refresh_output_label(self) -> None:
        if not self.output_section or not self.output_btn:
            return
        new_mode = self.mode_combo.currentText() == "Crear archivo nuevo"
        if new_mode:
            self.output_section.setText("Guardar resultado como")
            self.output_btn.setText("Elegir ruta")
            ph = "Sin resultado seleccionado"
        else:
            self.output_section.setText("Resultado existente")
            self.output_btn.setText("Seleccionar resultado")
            ph = "Sin resultado existente seleccionado"
        self._output_path = ""
        self.analysis = None
        self.path_lbl["output"].setText(ph)
        self.path_lbl["output"].setStyleSheet(_PATH_DEFAULT)

    def _set_path(self, key: str, value: str) -> None:
        display = value if len(value) <= 65 else "…" + value[-62:]
        self.path_lbl[key].setText(display)
        self.path_lbl[key].setStyleSheet(_PATH_ACTIVE)

    def _select_data(self) -> None:
        p, _ = QFileDialog.getOpenFileName(
            self, "Selecciona DATA mensual", "", "Excel (*.xlsx *.xlsm)")
        if p:
            self._data_path = p
            self._set_path("data", p)
            self.analysis = None

    def _select_output(self) -> None:
        if self.mode_combo.currentText() == "Crear archivo nuevo":
            p, _ = QFileDialog.getSaveFileName(
                self, "Guardar resultado como", "", "Excel (*.xlsx *.xlsm)")
        else:
            p, _ = QFileDialog.getOpenFileName(
                self, "Selecciona resultado existente", "", "Excel (*.xlsx *.xlsm)")
        if p:
            self._output_path = p
            self._set_path("output", p)

    def _select_folder(self) -> None:
        p = QFileDialog.getExistingDirectory(
            self, "Selecciona carpeta de expedientes")
        if p:
            self._folder_path = p
            self._set_path("folder", p)

    # ── Acciones ──────────────────────────────────────────────────────────────

    def _analyze(self) -> None:
        try:
            self._validate(require_output=False)
            month = self.month_combo.currentText()
            self.analysis = analyze_data_file(self._data_path, month)
            dup = f"\nLlaves duplicadas: {len(self.analysis.duplicate_keys)}" \
                  if self.analysis.duplicate_keys else ""
            self._log(f"DATA válido: {self.analysis.total_records} registros para {month}.{dup}")

            if self._output_path and Path(self._output_path).exists():
                has_sheet, has_rows = inspect_output_file(self._output_path, month)
                if has_sheet or has_rows:
                    self._log(f"Aviso: {month} ya existe en el resultado.")
        except ExcelValidationError as e:
            self._error(str(e))
        except Exception as e:
            self._error(f"Error inesperado al analizar: {e}")

    def _generate(self) -> None:
        try:
            self._validate(require_output=True)
            month = self.month_combo.currentText()
            if self.analysis is None:
                self.analysis = analyze_data_file(self._data_path, month)

            out = Path(self._output_path)
            has_sheet, has_rows = inspect_output_file(out, month)
            if has_sheet or has_rows or self.overwrite_chk.isChecked():
                reply = QMessageBox.question(
                    self, "Confirmar reescritura",
                    f"El mes {month} se reemplazará en el resultado.\n\n"
                    "Se creará un respaldo automático antes de modificar. ¿Continuar?",
                )
                if reply != QMessageBox.StandardButton.Yes:
                    self._log("Operación cancelada.")
                    return

            result = generate_result_file(
                output_path=out, analysis=self.analysis, month=month,
                overwrite_month=self.overwrite_chk.isChecked() or has_sheet or has_rows,
                expedients_folder=self._folder_path,
                create_backup=True,
            )
            backup = f"\nRespaldo: {result.backup_path}" if result.backup_path else ""
            self._log(
                f"Generado: {result.output_path}\n"
                f"Registros: {result.records_written}\n"
                f"PD encontrados: {result.pd_found}\n"
                f"PD faltantes: {result.pd_missing}\n"
                f"PS encontrados: {result.ps_found}\n"
                f"PS faltantes: {result.ps_missing}\n"
                f"XMLPD encontrados: {result.xmlpd_found}\n"
                f"XMLPD faltantes: {result.xmlpd_missing}\n"
                f"DODA/PITA/AVC encontrados: {result.doda_found}\n"
                f"DODA/PITA/AVC faltantes: {result.doda_missing}\n"
                f"DETALLE COVE encontrados: {result.detalle_cove_found}\n"
                f"DETALLE COVE faltantes: {result.detalle_cove_missing}\n"
                f"ACUSE COVE encontrados: {result.acuse_cove_found}\n"
                f"ACUSE COVE faltantes: {result.acuse_cove_missing}\n"
                f"Reemplazado: {'sí' if result.month_replaced else 'no'}{backup}"
            )
            QMessageBox.information(self, "Auditoría generada",
                                    "El resultado se generó correctamente.")
        except ExcelValidationError as e:
            self._error(str(e))
        except Exception as e:
            self._error(f"Error inesperado al generar: {e}")

    def _validate(self, require_output: bool) -> None:
        if not self._data_path:
            raise ExcelValidationError("Selecciona el archivo DATA mensual.")
        if require_output and not self._output_path:
            raise ExcelValidationError("Selecciona o define el archivo resultado.")
        if not self._folder_path:
            raise ExcelValidationError("Selecciona la carpeta de expedientes.")
        y = self.year_combo.currentText()
        if not y.isdigit() or len(y) != 4:
            raise ExcelValidationError("El año debe tener 4 dígitos.")

    def _log(self, message: str) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_lines.append(f"[{ts}] {message}")
        self.log_lines = self.log_lines[-14:]
        self.log_area.setPlainText("\n\n".join(self.log_lines))
        self.log_area.moveCursor(QTextCursor.MoveOperation.End)

    def _error(self, message: str) -> None:
        self._log(f"ERROR: {message}")
        QMessageBox.critical(self, "Error", message)


# ── Entry point ───────────────────────────────────────────────────────────────

def run_app() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    window = AuditorApp()
    window.show()
    app.exec()
