from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tkinter import BooleanVar, Button, Canvas, Checkbutton, OptionMenu, StringVar, Tk, filedialog, messagebox

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
            "border": "#334155",
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
        self.log_lines: list[str] = []

        self.path_text_items: dict[str, int] = {}
        self.output_label_item: int | None = None
        self.log_text_item: int | None = None

        self._build_layout()
        self._update_output_label()
        self._log("Listo. Selecciona el DATA mensual, resultado y carpeta de expedientes.")

    def _build_layout(self) -> None:
        self.canvas = Canvas(self, bg=self.colors["bg"], highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self._draw_background()
        self._draw_texts()
        self._create_controls()

    def _draw_background(self) -> None:
        c = self.canvas
        c.create_rectangle(0, 0, 1080, 720, fill=self.colors["bg"], outline="")
        c.create_rectangle(40, 116, 704, 660, fill=self.colors["card"], outline=self.colors["border"], width=1)
        c.create_rectangle(728, 116, 1040, 660, fill=self.colors["card"], outline=self.colors["border"], width=1)
        c.create_rectangle(68, 500, 676, 590, fill=self.colors["panel"], outline=self.colors["border"], width=1)
        c.create_rectangle(756, 174, 1012, 632, fill=self.colors["field"], outline=self.colors["border"], width=1)

    def _draw_texts(self) -> None:
        c = self.canvas
        c.create_text(40, 36, text="Auditor de Expedientes", fill=self.colors["text"], font=("Helvetica", 30, "bold"), anchor="nw")
        c.create_text(42, 78, text="Genera DS acumulada y hojas mensuales estrictas desde el DATA oficial.", fill=self.colors["muted"], font=("Helvetica", 14), anchor="nw")

        c.create_text(68, 144, text="Modo de operacion", fill=self.colors["label"], font=("Helvetica", 13, "bold"), anchor="nw")
        c.create_text(68, 214, text="Archivo DATA mensual", fill=self.colors["label"], font=("Helvetica", 13, "bold"), anchor="nw")
        self.output_label_item = c.create_text(68, 294, text="Guardar resultado como", fill=self.colors["label"], font=("Helvetica", 13, "bold"), anchor="nw")
        c.create_text(68, 374, text="Carpeta de expedientes", fill=self.colors["label"], font=("Helvetica", 13, "bold"), anchor="nw")
        c.create_text(68, 454, text="Mes", fill=self.colors["label"], font=("Helvetica", 13, "bold"), anchor="nw")
        c.create_text(348, 454, text="Año", fill=self.colors["label"], font=("Helvetica", 13, "bold"), anchor="nw")

        c.create_text(90, 520, text=f'El DATA debe contener una hoja llamada exactamente "{DATA_SHEET_NAME}".', fill=self.colors["cyan"], font=("Helvetica", 13, "bold"), anchor="nw")
        c.create_text(90, 548, text="Columnas obligatorias: Patente, Pedimento, SeccionAduanera, TipoOperacion, ClaveDocumento, FechaPagoReal.", fill=self.colors["label"], font=("Helvetica", 12), anchor="nw", width=560)

        c.create_text(756, 144, text="Estado", fill=self.colors["text"], font=("Helvetica", 20, "bold"), anchor="nw")
        self.log_text_item = c.create_text(776, 194, text="", fill="#dbeafe", font=("Menlo", 12), anchor="nw", width=216)

        self.path_text_items["data"] = self._path_box(68, 244, "Sin DATA seleccionado")
        self.path_text_items["output"] = self._path_box(68, 324, "Sin resultado seleccionado")
        self.path_text_items["folder"] = self._path_box(68, 404, "Sin carpeta seleccionada")

    def _path_box(self, x: int, y: int, text: str) -> int:
        self.canvas.create_rectangle(x, y, x + 410, y + 34, fill=self.colors["field"], outline=self.colors["border"], width=1)
        return self.canvas.create_text(x + 12, y + 9, text=text, fill=self.colors["muted"], font=("Helvetica", 11), anchor="nw", width=386)

    def _create_controls(self) -> None:
        mode = OptionMenu(self.canvas, self.mode_var, "Crear archivo nuevo", "Actualizar archivo existente", command=lambda _value: self._update_output_label())
        self._style_option_menu(mode)
        self.canvas.create_window(250, 154, window=mode, width=330, height=38, anchor="nw")

        self.canvas.create_window(504, 242, window=self._button("Seleccionar DATA", self._select_data_file, self.colors["neutral"]), width=160, height=38, anchor="nw")
        self.output_button = self._button("Elegir ruta", self._select_output_file, self.colors["neutral"])
        self.canvas.create_window(504, 322, window=self.output_button, width=160, height=38, anchor="nw")
        self.canvas.create_window(504, 402, window=self._button("Seleccionar carpeta", self._select_folder, self.colors["neutral"]), width=160, height=38, anchor="nw")

        month = OptionMenu(self.canvas, self.month_var, *MONTHS)
        self._style_option_menu(month)
        self.canvas.create_window(68, 478, window=month, width=230, height=38, anchor="nw")

        years = [str(datetime.now().year + offset) for offset in range(-2, 4)]
        year = OptionMenu(self.canvas, self.year_var, *years)
        self._style_option_menu(year)
        self.canvas.create_window(348, 478, window=year, width=130, height=38, anchor="nw")

        rewrite = Checkbutton(
            self.canvas,
            text="Permitir reescribir el mes con confirmacion",
            variable=self.overwrite_var,
            bg=self.colors["card"],
            fg=self.colors["label"],
            activebackground=self.colors["card"],
            activeforeground=self.colors["cyan"],
            selectcolor=self.colors["field"],
            font=("Helvetica", 12),
        )
        self.canvas.create_window(68, 610, window=rewrite, anchor="nw")

        self.canvas.create_window(68, 676, window=self._button("Analizar", self._analyze, self.colors["blue"]), width=290, height=42, anchor="sw")
        self.canvas.create_window(386, 676, window=self._button("Generar Auditoria", self._generate, self.colors["accent"]), width=290, height=42, anchor="sw")

    def _button(self, text: str, command, bg: str) -> Button:
        return Button(
            self.canvas,
            text=text,
            command=command,
            bg=bg,
            fg="#ffffff",
            activebackground=bg,
            activeforeground="#ffffff",
            relief="flat",
            padx=14,
            pady=8,
            font=("Helvetica", 12, "bold"),
            cursor="hand2",
        )

    def _style_option_menu(self, widget: OptionMenu) -> None:
        widget.configure(
            bg=self.colors["field"],
            fg=self.colors["text"],
            activebackground=self.colors["neutral"],
            activeforeground=self.colors["text"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            font=("Helvetica", 12, "bold"),
        )
        widget["menu"].configure(bg=self.colors["field"], fg=self.colors["text"], activebackground=self.colors["neutral"])

    def _first_render(self) -> None:
        self.update_idletasks()
        self.lift()

    def _set_path_text(self, key: str, value: str) -> None:
        display = value if len(value) <= 58 else "..." + value[-55:]
        self.canvas.itemconfigure(self.path_text_items[key], text=display, fill=self.colors["text"])

    def _update_output_label(self) -> None:
        if self.output_label_item is None:
            return
        if self.mode_var.get() == "Crear archivo nuevo":
            self.canvas.itemconfigure(self.output_label_item, text="Guardar resultado como")
            self.output_button.configure(text="Elegir ruta")
            placeholder = "Sin resultado seleccionado"
        else:
            self.canvas.itemconfigure(self.output_label_item, text="Resultado existente")
            self.output_button.configure(text="Seleccionar resultado")
            placeholder = "Sin resultado existente seleccionado"
        self.output_path_var.set("")
        self.analysis = None
        if "output" in self.path_text_items:
            self.canvas.itemconfigure(self.path_text_items["output"], text=placeholder, fill=self.colors["muted"])

    def _select_data_file(self) -> None:
        path = filedialog.askopenfilename(title="Selecciona DATA mensual", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            self.data_path_var.set(path)
            self._set_path_text("data", path)
            self.analysis = None

    def _select_output_file(self) -> None:
        if self.mode_var.get() == "Crear archivo nuevo":
            path = filedialog.asksaveasfilename(title="Guardar resultado como", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx *.xlsm")])
        else:
            path = filedialog.askopenfilename(title="Selecciona resultado existente", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            self.output_path_var.set(path)
            self._set_path_text("output", path)

    def _select_folder(self) -> None:
        path = filedialog.askdirectory(title="Selecciona carpeta de expedientes")
        if path:
            self.folder_path_var.set(path)
            self._set_path_text("folder", path)

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
        self.log_lines.append(f"[{timestamp}] {message}")
        self.log_lines = self.log_lines[-14:]
        if self.log_text_item is not None:
            self.canvas.itemconfigure(self.log_text_item, text="\n\n".join(self.log_lines))

    def _error(self, message: str) -> None:
        self._log(f"ERROR: {message}")
        messagebox.showerror("Error", message)


def run_app() -> None:
    app = AuditorApp()
    app.mainloop()
