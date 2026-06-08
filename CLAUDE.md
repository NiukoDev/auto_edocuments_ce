# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Instalar dependencias
source .venv/bin/activate
pip install -r requirements.txt

# Ejecutar la app
python src/main.py

# Generar ejecutable (solo Windows)
pyinstaller --noconsole --onefile --name AuditorExpedientes src/main.py
```

> El archivo `src/main.py` debe ejecutarse desde `src/` como directorio de trabajo o con `python src/main.py` desde la raíz, ya que los imports son relativos al package `src` (sin prefijo de módulo).

No hay suite de pruebas automatizadas.

## Arquitectura

Aplicación de escritorio (PySide6 / Qt) que procesa archivos Excel mensuales de datos aduaneros y genera un archivo de auditoría anual acumulativo.

### Flujo principal

1. El usuario selecciona un **archivo DATA mensual** (`.xlsx`/`.xlsm`) con hoja `Datos Generales`.
2. El usuario selecciona el **archivo resultado** (nuevo o existente) y, opcionalmente, una **carpeta de expedientes**.
3. **Analizar**: `excel_service.analyze_data_file` lee el DATA y produce un `AnalysisResult` con la lista de `DataRecord`.
4. **Generar Auditoría**: `excel_service.generate_result_file` escribe en el archivo resultado:
   - Hoja `DS` (acumulativa anual): registros de todos los meses.
   - Hoja del mes (p. ej. `ENERO`): tabla con columnas de documentos, ordenada por `TipoOperacion` + `ClaveDocumento`.

### Módulos

| Archivo | Responsabilidad |
|---|---|
| `ui.py` | `AuditorApp` (QMainWindow) — toda la UI y orquestación de eventos |
| `excel_service.py` | Lógica de negocio: leer DATA, generar/actualizar resultado, estilos Excel |
| `file_scanner.py` | Construye índice de archivos (recursivo, lowercase) y verifica presencia de documentos |
| `models.py` | Dataclasses inmutables: `DataRecord`, `AnalysisResult`, `GenerationResult` |
| `config.py` | Todas las constantes: columnas, reglas de documentos, patentes, meses |

### Validación de documentos

La clave de expediente (`llave`) tiene el formato `{SeccionAduanera}-{Patente}-{Pedimento}`.

`file_scanner.has_document` busca en el índice de archivos si hay un fichero cuyo nombre contenga el patrón de documento **y** la `llave` compacta (sin guiones) o el número de pedimento.

Las reglas de búsqueda viven en `config.py`:

- `DOCUMENT_RULES`: patrón y extensión por tipo de documento (regla general).
- `DOCUMENT_PATTERNS_BY_PATENT`: patrones alternativos por patente (`3183`, `3685`). `excel_service._document_patterns` combina el patrón general con los específicos.
- `DOCUMENTS_NOT_APPLICABLE_BY_COLUMN`: condiciones `TipoOperacion` + `ClaveDocumento` que marcan un documento como `NA` en lugar de buscarlo.

### Gestión del archivo resultado

- Si el mes ya existe (hoja o filas en DS), se pide confirmación antes de reemplazar.
- Antes de modificar un archivo existente, se crea un backup automático con timestamp.
- Las hojas se reordenan para que `DS` quede siempre primera.
