from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from config import APP_NAME, DATA_SHEET_NAME, MONTHS
from excel_service import ExcelValidationError, analyze_data_file, generate_result_file, inspect_output_file
from models import AnalysisResult


class AuditorApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title(APP_NAME)
        self.geometry("1080x720")
        self.minsize(920, 640)

        self.mode_var = ctk.StringVar(value="Crear archivo nuevo")
        self.data_path_var = ctk.StringVar()
        self.output_path_var = ctk.StringVar()
        self.folder_path_var = ctk.StringVar()
        self.month_var = ctk.StringVar(value=MONTHS[datetime.now().month - 1])
        self.year_var = ctk.StringVar(value=str(datetime.now().year))
        self.overwrite_var = ctk.BooleanVar(value=False)
        self.analysis: AnalysisResult | None = None

        self._build_layout()
        self._update_output_label()

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        root = ctk.CTkFrame(self, fg_color="#0b1120")
        root.grid(row=0, column=0, sticky="nsew")
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(root, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", padx=26, pady=(22, 10))
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            header,
            text="Auditor de Expedientes",
            font=ctk.CTkFont(size=30, weight="bold"),
            text_color="#e5f0ff",
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = ctk.CTkLabel(
            header,
            text="Genera DS acumulada y hojas mensuales estrictas desde el DATA oficial.",
            font=ctk.CTkFont(size=14),
            text_color="#94a3b8",
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(3, 0))

        form = ctk.CTkFrame(root, fg_color="#111827", corner_radius=18)
        form.grid(row=1, column=0, sticky="nsew", padx=(26, 12), pady=(10, 26))
        form.grid_columnconfigure(1, weight=1)

        self._add_mode_selector(form, 0)
        self._add_file_row(form, 1, "DATA mensual", self.data_path_var, self._select_data_file, "Seleccionar DATA")
        self.output_label = ctk.CTkLabel(form, text="Resultado", text_color="#cbd5e1", font=ctk.CTkFont(size=14, weight="bold"))
        self.output_label.grid(row=2, column=0, sticky="w", padx=18, pady=(16, 4))
        output_entry = ctk.CTkEntry(form, textvariable=self.output_path_var, height=38)
        output_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=(16, 4))
        self.output_button = ctk.CTkButton(form, text="Seleccionar", command=self._select_output_file, height=38)
        self.output_button.grid(row=2, column=2, sticky="ew", padx=(8, 18), pady=(16, 4))

        self._add_file_row(form, 3, "Carpeta expedientes", self.folder_path_var, self._select_folder, "Seleccionar carpeta")

        period = ctk.CTkFrame(form, fg_color="transparent")
        period.grid(row=4, column=0, columnspan=3, sticky="ew", padx=18, pady=(18, 4))
        period.grid_columnconfigure(1, weight=1)
        period.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(period, text="Mes", text_color="#cbd5e1", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkOptionMenu(period, variable=self.month_var, values=MONTHS, height=38).grid(row=0, column=1, sticky="ew", padx=(10, 18))
        ctk.CTkLabel(period, text="Año", text_color="#cbd5e1", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=2, sticky="w")
        ctk.CTkEntry(period, textvariable=self.year_var, height=38, width=110).grid(row=0, column=3, sticky="ew", padx=(10, 0))

        overwrite = ctk.CTkSwitch(
            form,
            text="Permitir reescribir el mes con confirmacion",
            variable=self.overwrite_var,
            text_color="#cbd5e1",
            progress_color="#22d3ee",
        )
        overwrite.grid(row=5, column=0, columnspan=3, sticky="w", padx=18, pady=(16, 8))

        instructions = ctk.CTkFrame(form, fg_color="#172033", corner_radius=14)
        instructions.grid(row=6, column=0, columnspan=3, sticky="ew", padx=18, pady=(12, 14))
        instructions.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            instructions,
            text=f'El DATA debe contener una hoja llamada exactamente "{DATA_SHEET_NAME}".',
            text_color="#67e8f9",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 4))
        ctk.CTkLabel(
            instructions,
            text="Columnas obligatorias: Patente, Pedimento, SeccionAduanera, TipoOperacion, ClaveDocumento, FechaPagoReal.",
            text_color="#cbd5e1",
            wraplength=560,
            justify="left",
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 12))

        buttons = ctk.CTkFrame(form, fg_color="transparent")
        buttons.grid(row=7, column=0, columnspan=3, sticky="ew", padx=18, pady=(4, 18))
        buttons.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(buttons, text="Analizar", command=self._analyze, height=44, fg_color="#2563eb").grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(buttons, text="Generar Auditoria", command=self._generate, height=44, fg_color="#0891b2").grid(row=0, column=1, sticky="ew", padx=(8, 0))

        side = ctk.CTkFrame(root, fg_color="#111827", corner_radius=18)
        side.grid(row=1, column=1, sticky="nsew", padx=(12, 26), pady=(10, 26))
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(side, text="Estado", font=ctk.CTkFont(size=20, weight="bold"), text_color="#e5f0ff").grid(
            row=0, column=0, sticky="w", padx=18, pady=(18, 8)
        )
        self.log_box = ctk.CTkTextbox(side, fg_color="#020617", text_color="#dbeafe", corner_radius=14, wrap="word")
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))
        self._log("Listo. Selecciona el DATA mensual, resultado y carpeta de expedientes.")

    def _add_mode_selector(self, parent: ctk.CTkFrame, row: int) -> None:
        ctk.CTkLabel(parent, text="Modo", text_color="#cbd5e1", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=row, column=0, sticky="w", padx=18, pady=(18, 4)
        )
        segmented = ctk.CTkSegmentedButton(
            parent,
            values=["Crear archivo nuevo", "Actualizar archivo existente"],
            variable=self.mode_var,
            command=lambda _: self._update_output_label(),
            height=38,
        )
        segmented.grid(row=row, column=1, columnspan=2, sticky="ew", padx=(8, 18), pady=(18, 4))

    def _add_file_row(self, parent, row: int, label: str, variable: ctk.StringVar, command, button_text: str) -> None:
        ctk.CTkLabel(parent, text=label, text_color="#cbd5e1", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=row, column=0, sticky="w", padx=18, pady=(16, 4)
        )
        ctk.CTkEntry(parent, textvariable=variable, height=38).grid(row=row, column=1, sticky="ew", padx=8, pady=(16, 4))
        ctk.CTkButton(parent, text=button_text, command=command, height=38).grid(row=row, column=2, sticky="ew", padx=(8, 18), pady=(16, 4))

    def _update_output_label(self) -> None:
        if not hasattr(self, "output_label"):
            return
        if self.mode_var.get() == "Crear archivo nuevo":
            self.output_label.configure(text="Guardar resultado como")
            self.output_button.configure(text="Elegir ruta")
        else:
            self.output_label.configure(text="Resultado existente")
            self.output_button.configure(text="Seleccionar resultado")
        self.output_path_var.set("")
        self.analysis = None

    def _select_data_file(self) -> None:
        path = filedialog.askopenfilename(title="Selecciona DATA mensual", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            self.data_path_var.set(path)
            self.analysis = None

    def _select_output_file(self) -> None:
        if self.mode_var.get() == "Crear archivo nuevo":
            path = filedialog.asksaveasfilename(
                title="Guardar resultado como",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx *.xlsm")],
            )
        else:
            path = filedialog.askopenfilename(title="Selecciona resultado existente", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            self.output_path_var.set(path)

    def _select_folder(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de expedientes")
        if path:
            self.folder_path_var.set(path)

    def _analyze(self) -> None:
        try:
            self._validate_form(require_output=False)
            month = self.month_var.get()
            self.analysis = analyze_data_file(self.data_path_var.get(), month)
            duplicate_msg = ""
            if self.analysis.duplicate_keys:
                duplicate_msg = f"\nLlaves duplicadas en DATA: {len(self.analysis.duplicate_keys)}"
            self._log(f"DATA valido: {self.analysis.total_records} registros para {month}.{duplicate_msg}")

            output_path = self.output_path_var.get().strip()
            if output_path and Path(output_path).exists():
                has_sheet, has_rows = inspect_output_file(output_path, month)
                if has_sheet or has_rows:
                    self._log(f"Aviso: {month} ya existe en el resultado. Para reemplazarlo activa la reescritura.")
        except ExcelValidationError as exc:
            self._error(str(exc))
        except Exception as exc:
            self._error(f"Error inesperado al analizar: {exc}")

    def _generate(self) -> None:
        try:
            self._validate_form(require_output=True)
            month = self.month_var.get()
            if self.analysis is None:
                self.analysis = analyze_data_file(self.data_path_var.get(), month)

            output_path = Path(self.output_path_var.get())
            has_sheet, has_rows = inspect_output_file(output_path, month)
            needs_confirmation = has_sheet or has_rows or self.overwrite_var.get()
            if needs_confirmation:
                confirmed = messagebox.askyesno(
                    "Confirmar reescritura",
                    f"El mes {month} se reemplazara en el resultado.\n\n"
                    f"Se actualizara la hoja {month} y se eliminaran de DS las filas de {month} antes de insertar las nuevas.\n\n"
                    "Se creara un respaldo automatico antes de modificar. ¿Deseas continuar?",
                )
                if not confirmed:
                    self._log("Operacion cancelada por el usuario.")
                    return

            result = generate_result_file(
                output_path=output_path,
                analysis=self.analysis,
                month=month,
                overwrite_month=self.overwrite_var.get() or has_sheet or has_rows,
                create_backup=True,
            )
            backup_msg = f"\nRespaldo: {result.backup_path}" if result.backup_path else ""
            self._log(
                f"Resultado generado: {result.output_path}\n"
                f"Registros escritos: {result.records_written}\n"
                f"Mes reemplazado: {'si' if result.month_replaced else 'no'}{backup_msg}"
            )
            messagebox.showinfo("Auditoria generada", "El resultado se genero correctamente.")
        except ExcelValidationError as exc:
            self._error(str(exc))
        except Exception as exc:
            self._error(f"Error inesperado al generar: {exc}")

    def _validate_form(self, require_output: bool) -> None:
        if not self.data_path_var.get().strip():
            raise ExcelValidationError("Selecciona el archivo DATA mensual.")
        if require_output and not self.output_path_var.get().strip():
            raise ExcelValidationError("Selecciona o define el archivo resultado.")
        if not self.folder_path_var.get().strip():
            raise ExcelValidationError("Selecciona la carpeta de expedientes. En esta version aun no se audita, pero se requiere para mantener el flujo.")
        if not self.year_var.get().strip().isdigit() or len(self.year_var.get().strip()) != 4:
            raise ExcelValidationError("El año debe tener 4 digitos.")

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", f"[{timestamp}] {message}\n\n")
        self.log_box.see("end")

    def _error(self, message: str) -> None:
        self._log(f"ERROR: {message}")
        messagebox.showerror("Error", message)


def run_app() -> None:
    app = AuditorApp()
    app.mainloop()
