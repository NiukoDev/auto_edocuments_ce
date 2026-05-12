from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tkinter import BooleanVar, StringVar, Text, Tk, filedialog, messagebox, ttk

from config import APP_NAME, DATA_SHEET_NAME, MONTHS
from excel_service import ExcelValidationError, analyze_data_file, generate_result_file, inspect_output_file
from models import AnalysisResult


class AuditorApp(Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title(APP_NAME)
        self.geometry("1080x720")
        self.minsize(920, 640)
        self.configure(bg="#0b1120")

        self.mode_var = StringVar(value="Crear archivo nuevo")
        self.data_path_var = StringVar()
        self.output_path_var = StringVar()
        self.folder_path_var = StringVar()
        self.month_var = StringVar(value=MONTHS[datetime.now().month - 1])
        self.year_var = StringVar(value=str(datetime.now().year))
        self.overwrite_var = BooleanVar(value=False)
        self.analysis: AnalysisResult | None = None

        self._configure_style()
        self._build_layout()
        self._update_output_label()
        self.after(50, self._first_render)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("App.TFrame", background="#0b1120")
        style.configure("Card.TFrame", background="#111827", relief="flat")
        style.configure("Info.TFrame", background="#172033", relief="flat")
        style.configure("Title.TLabel", background="#0b1120", foreground="#e5f0ff", font=("Helvetica", 30, "bold"))
        style.configure("Subtitle.TLabel", background="#0b1120", foreground="#94a3b8", font=("Helvetica", 14))
        style.configure("Label.TLabel", background="#111827", foreground="#cbd5e1", font=("Helvetica", 13, "bold"))
        style.configure("InfoTitle.TLabel", background="#172033", foreground="#67e8f9", font=("Helvetica", 13, "bold"))
        style.configure("InfoText.TLabel", background="#172033", foreground="#cbd5e1", font=("Helvetica", 12))
        style.configure("StatusTitle.TLabel", background="#111827", foreground="#e5f0ff", font=("Helvetica", 20, "bold"))
        style.configure("TEntry", fieldbackground="#020617", foreground="#e5f0ff", insertcolor="#e5f0ff")
        style.configure("TCombobox", fieldbackground="#020617", background="#020617", foreground="#e5f0ff")
        style.configure("TCheckbutton", background="#111827", foreground="#cbd5e1", font=("Helvetica", 12))
        style.map("TCheckbutton", background=[("active", "#111827")], foreground=[("active", "#67e8f9")])
        style.configure("Primary.TButton", background="#2563eb", foreground="#ffffff", font=("Helvetica", 12, "bold"), padding=10)
        style.configure("Accent.TButton", background="#0891b2", foreground="#ffffff", font=("Helvetica", 12, "bold"), padding=10)
        style.configure("Neutral.TButton", background="#1f2937", foreground="#e5f0ff", font=("Helvetica", 11), padding=8)
        style.map("Primary.TButton", background=[("active", "#1d4ed8")])
        style.map("Accent.TButton", background=[("active", "#0e7490")])
        style.map("Neutral.TButton", background=[("active", "#374151")])

    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, style="App.TFrame", padding=26)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=2)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(1, weight=1)

        header = ttk.Frame(root, style="App.TFrame")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Auditor de Expedientes", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Genera DS acumulada y hojas mensuales estrictas desde el DATA oficial.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))

        form = ttk.Frame(root, style="Card.TFrame", padding=18)
        form.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Modo", style="Label.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        mode = ttk.Combobox(form, textvariable=self.mode_var, values=["Crear archivo nuevo", "Actualizar archivo existente"], state="readonly")
        mode.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(10, 0), pady=(0, 8))
        mode.bind("<<ComboboxSelected>>", lambda _event: self._update_output_label())

        self._add_file_row(form, 1, "DATA mensual", self.data_path_var, self._select_data_file, "Seleccionar DATA")

        self.output_label = ttk.Label(form, text="Resultado", style="Label.TLabel")
        self.output_label.grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(form, textvariable=self.output_path_var).grid(row=2, column=1, sticky="ew", padx=(10, 10), pady=8)
        self.output_button = ttk.Button(form, text="Seleccionar", command=self._select_output_file, style="Neutral.TButton")
        self.output_button.grid(row=2, column=2, sticky="ew", pady=8)

        self._add_file_row(form, 3, "Carpeta expedientes", self.folder_path_var, self._select_folder, "Seleccionar carpeta")

        ttk.Label(form, text="Mes", style="Label.TLabel").grid(row=4, column=0, sticky="w", pady=8)
        ttk.Combobox(form, textvariable=self.month_var, values=MONTHS, state="readonly").grid(row=4, column=1, sticky="ew", padx=(10, 10), pady=8)
        ttk.Entry(form, textvariable=self.year_var, width=10).grid(row=4, column=2, sticky="ew", pady=8)

        ttk.Checkbutton(
            form,
            text="Permitir reescribir el mes con confirmacion",
            variable=self.overwrite_var,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(14, 12))

        instructions = ttk.Frame(form, style="Info.TFrame", padding=14)
        instructions.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(4, 14))
        instructions.columnconfigure(0, weight=1)
        ttk.Label(
            instructions,
            text=f'El DATA debe contener una hoja llamada exactamente "{DATA_SHEET_NAME}".',
            style="InfoTitle.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            instructions,
            text="Columnas obligatorias: Patente, Pedimento, SeccionAduanera, TipoOperacion, ClaveDocumento, FechaPagoReal.",
            style="InfoText.TLabel",
            wraplength=590,
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        buttons = ttk.Frame(form, style="Card.TFrame")
        buttons.grid(row=7, column=0, columnspan=3, sticky="ew")
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)
        ttk.Button(buttons, text="Analizar", command=self._analyze, style="Primary.TButton").grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(buttons, text="Generar Auditoria", command=self._generate, style="Accent.TButton").grid(row=0, column=1, sticky="ew", padx=(8, 0))

        side = ttk.Frame(root, style="Card.TFrame", padding=18)
        side.grid(row=1, column=1, sticky="nsew", padx=(12, 0))
        side.columnconfigure(0, weight=1)
        side.rowconfigure(1, weight=1)
        ttk.Label(side, text="Estado", style="StatusTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 12))
        self.log_box = Text(
            side,
            bg="#020617",
            fg="#dbeafe",
            insertbackground="#dbeafe",
            relief="flat",
            wrap="word",
            font=("Menlo", 12),
            padx=12,
            pady=12,
        )
        self.log_box.grid(row=1, column=0, sticky="nsew")
        self._log("Listo. Selecciona el DATA mensual, resultado y carpeta de expedientes.")

    def _add_file_row(self, parent, row: int, label: str, variable: StringVar, command, button_text: str) -> None:
        ttk.Label(parent, text=label, style="Label.TLabel").grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(10, 10), pady=8)
        ttk.Button(parent, text=button_text, command=command, style="Neutral.TButton").grid(row=row, column=2, sticky="ew", pady=8)

    def _first_render(self) -> None:
        self.update_idletasks()
        self.lift()

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
