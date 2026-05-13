# AGENTS.md

## Setup y comandos
- Crear entorno en macOS/Linux: `python3 -m venv .venv && source .venv/bin/activate && python -m pip install --upgrade pip && pip install -r requirements.txt`.
- Crear entorno en Windows: `py -3.12 -m venv .venv`, luego `.venv\Scripts\activate`, `python -m pip install --upgrade pip`, `pip install -r requirements.txt`.
- Ejecutar app local: `python src/main.py` o `.venv/bin/python src/main.py` desde la raiz del repo.
- Verificacion rapida disponible: `.venv/bin/python -m compileall src`.
- Generar ejecutable Windows desde Windows/GitHub Actions runner, no desde macOS: `pyinstaller --noconsole --onefile --name AuditorExpedientes src/main.py`.

## Arquitectura
- Entrada de escritorio: `src/main.py`; importa `run_app` desde la UI.
- UI: `src/ui.py`; usa `PySide6` porque el Tk de macOS usado durante desarrollo renderizo mal widgets nativos (`Label`/`Entry`) y tambien requirio intentos con `Canvas`.
- Logica Excel: `src/excel_service.py`; ahi viven lectura estricta del DATA, generacion de `DS`, respaldo y reescritura mensual.
- Constantes de negocio: `src/config.py`; cambia nombres de hojas/columnas ahi antes de tocar servicios.
- Modelos simples: `src/models.py`.

## Reglas de negocio verificadas
- El DATA debe tener hoja llamada exactamente `Datos Generales`; no agregar tolerancia a variantes salvo instruccion explicita.
- Columnas obligatorias del DATA: `Patente`, `Pedimento`, `SeccionAduanera`, `TipoOperacion`, `ClaveDocumento`, `FechaPagoReal`.
- La llave se recalcula como `SeccionAduanera-Patente-Pedimento`; no confiar en una columna `LLAVE` del DATA.
- `DS` es acumulada anual. Al reescribir un mes, eliminar primero las filas de ese `MES` en `DS` y recrear la hoja mensual.
- La app aun no valida documentos de la carpeta de expedientes; por ahora crea columnas de auditoria y deja `ESTATUS = PENDIENTE`.

## Datos locales y repo
- `doc_aux/` contiene muestras/expedientes locales y esta ignorado; no lo versionar.
- Archivos temporales de Excel `~$*.xls*` estan ignorados y tambien se rechazan en validacion.
- No hay tests, CI ni workflow de GitHub Actions configurados todavia.
