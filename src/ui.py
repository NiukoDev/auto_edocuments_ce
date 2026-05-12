from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tkinter import BooleanVar, Button, Checkbutton, Entry, Frame, Label, OptionMenu, StringVar, Text, Tk, filedialog, messagebox

from config import APP_NAME, DATA_SHEET_NAME, MONTHS
from excel_service import ExcelValidationError, analyze_data_file, generate_result_file, inspect_output_file
from models import AnalysisResult


class AuditorApp(Tk):
    def __init__(self) -> None:
        super().__init__()

        self.colors = {
            "bg": "#0b1120",
            "card": "#111827",
            "panel": "#172033",
            "field": "#020617",
            "text": "#e5f0ff",
            "muted": "#94a3b8",
            "label": "#cbd5e1",
            "cyan": "#67e8f9",
            "blue": "#2563eb",
            "accent": "#0891b2",
            "neutral": "#1f2937",
        }

        self.title(APP_NAME)
        self.geometry("1080x720")
        self.minsize(920, 640)
        self.configure(bg=self.colors["bg"])

        self.mode_var = StringVar(value="Crear archivo nuevo")
        self.data_path_var = StringVar()
        self.output_path_var = StringVar()
        self.folder_path_var = StringVar()
        self.month_var = StringVar(value=MONTHS[datetime.now().month - 1])
        self.year_var = StringVar(value=str(datetime.now().year))
        self.overwrite_var = BooleanVar(value=False)
        self.analysis: AnalysisResult | None = None

        self._build_layout()
        self._update_output_label()
        self.after(50, self._first_render)

    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        root = Frame(self, bg=self.colors["bg"], padx=26, pady=26)
        root.grid(row=0, column=0, sticky="nsew")
        root.grid_columnconfigure(0, weight=2)
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(1, weight=1)

        header = Frame(root, bg=self.colors["bg"])
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 16))
        header.grid_columnconfigure(0, weight=1)
        Label(header, text="Auditor de Expedientes", bg=self.colors["bg"], fg=self.colors["text"], font=("Helvetica", 30, "bold"), anchor="w").grid(row=0, column=0, sticky="ew")
        Label(header, text="Genera DS acumulada y hojas mensuales estrictas desde el DATA oficial.", bg=self.colors["bg"], fg=self.colors["muted"], font=("Helvetica", 14), anchor="w").grid(row=1, column=0, sticky="ew", pady=(3, 0))

        form = Frame(root, bg=self.colors["card"], padx=18, pady=18)
        form.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        form.grid_columnconfigure(1, weight=1)

        self._label(form, "Modo").grid(row=0, column=0, sticky="w", pady=(0, 8))
        mode = OptionMenu(form, self.mode_var, "Crear archivo nuevo", "Actualizar archivo existente", command=lambda _value: self._update_output_label())
        self._style_option_menu(mode)
        mode.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(10, 0), pady=(0, 8))

        self._add_file_row(form, 1, "DATA mensual", self.data_path_var, self._select_data_file, "Seleccionar DATA")

        self.output_label = self._label(form, "Resultado")
        self.output_label.grid(row=2, column=0, sticky="w", pady=8)
        self._entry(form, self.output_path_var).grid(row=2, column=1, sticky="ew", padx=(10, 10), pady=8)
        self.output_button = self._button(form, "Seleccionar", self._select_output_file, self.colors["neutral"])
        self.output_button.grid(row=2, column=2, sticky="ew", pady=8)

        self._add_file_row(form, 3, "Carpeta expedientes", self.folder_path_var, self._select_folder, "Seleccionar carpeta")

        self._label(form, "Mes").grid(row=4, column=0, sticky="w", pady=8)
        month = OptionMenu(form, self.month_var, *MONTHS)
        self._style_option_menu(month)
        month.grid(row=4, column=1, sticky="ew", padx=(10, 10), pady=8)
        self._entry(form, self.year_var).grid(row=4, column=2, sticky="ew", pady=8)

        Checkbutton(
            form,
            text="Permitir reescribir el mes con confirmacion",
            variable=self.overwrite_var,
            bg=self.colors["card"],
            fg=self.colors["label"],
            activebackground=self.colors["card"],
            activeforeground=self.colors["cyan"],
            selectcolor=self.colors["field"],
            font=("Helvetica", 12),
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(14, 12))

        instructions = Frame(form, bg=self.colors["panel"], padx=14, pady=14)
        instructions.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(4, 14))
        instructions.grid_columnconfigure(0, weight=1)
        Label(instructions, text=f'El DATA debe contener una hoja llamada exactamente "{DATA_SHEET_NAME}".', bg=self.colors["panel"], fg=self.colors["cyan"], font=("Helvetica", 13, "bold"), anchor="w").grid(row=0, column=0, sticky="ew")
        Label(instructions, text="Columnas obligatorias: Patente, Pedimento, SeccionAduanera, TipoOperacion, ClaveDocumento, FechaPagoReal.", bg=self.colors["panel"], fg=self.colors["label"], font=("Helvetica", 12), anchor="w", justify="left", wraplength=590).grid(row=1, column=0, sticky="ew", pady=(6, 0))

        buttons = Frame(form, bg=self.colors["card"])
        buttons.grid(row=7, column=0, columnspan=3, sticky="ew")
        buttons.grid_columnconfigure(0, weight=1)
        buttons.grid_columnconfigure(1, weight=1)
        self._button(buttons, "Analizar", self._analyze, self.colors["blue"]).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._button(buttons, "Generar Auditoria", self._generate, self.colors["accent"]).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        side = Frame(root, bg=self.colors["card"], padx=18, pady=18)
        side.grid(row=1, column=1, sticky="nsew", padx=(12, 0))
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(1, weight=1)
        Label(side, text="Estado", bg=self.colors["card"], fg=self.colors["text"], font=("Helvetica", 20, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", pady=(0, 12))
        self.log_box = Text(side, bg=self.colors["field"], fg="#dbeafe", insertbackground="#dbeafe", relief="flat", wrap="word", font=("Menlo", 12), padx=12, pady=12)
        self.log_box.grid(row=1, column=0, sticky="nsew")
        self._log("Listo. Selecciona el DATA mensual, resultado y carpeta de expedientes.")

    def _label(self, parent, text: str) -> Label:
        return Label(parent, text=text, bg=self.colors["card"], fg=self.colors["label"], font=("Helvetica", 13, "bold"), anchor="w")

    def _entry(self, parent, variable: StringVar) -> Entry:
        return Entry(parent, textvariable=variable, bg=self.colors["field"], fg=self.colors["text"], insertbackground=self.colors["text"], relief="flat", highlightthickness=1, highlightbackground="#334155", highlightcolor=self.colors["cyan"], font=("Helvetica", 12))

    def _button(self, parent, text: str, command, bg: str) -> Button:
        return Button(parent, text=text, command=command, bg=bg, fg="#ffffff", activebackground=bg, activeforeground="#ffffff", relief="flat", padx=14, pady=9, font=("Helvetica", 12, "bold"), cursor="hand2")

    def _style_option_menu(self, widget: OptionMenu) -> None:
        widget.configure(bg=self.colors["field"], fg=self.colors["text"], activebackground=self.colors["neutral"], activeforeground=self.colors["text"], relief="flat", highlightthickness=1, highlightbackground="#334155", font=("Helvetica", 12))
        widget["menu"].configure(bg=self.colors["field"], fg=self.colors["text"], activebackground=self.colors["neutral"])

    def _add_file_row(self, parent, row: int, label: str, variable: StringVar, command, button_text: str) -> None:
        self._label(parent, label).grid(row=row, column=0, sticky="w", pady=8)
        self._entry(parent, variable).grid(row=row, column=1, sticky="ew", padx=(10, 10), pady=8)
        self._button(parent, button_text, command, self.colors["neutral"]).grid(row=row, column=2, sticky="ew", pady=8)

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
            path = filedialog.asksaveasfilename(title="Guardar resultado como", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx *.xlsm")])
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

            result = generate_result_file(output_path=output_path, analysis=self.analysis, month=month, overwrite_month=self.overwrite_var.get() or has_sheet or has_rows, create_backup=True)
            backup_msg = f"\nRespaldo: {result.backup_path}" if result.backup_path else ""
            self._log(f"Resultado generado: {result.output_path}\nRegistros escritos: {result.records_written}\nMes reemplazado: {'si' if result.month_replaced else 'no'}{backup_msg}")
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
