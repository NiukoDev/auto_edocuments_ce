# Auditor de Expedientes

Aplicacion de escritorio para generar un archivo anual de auditoria a partir del DATA mensual.

## Requisitos del DATA

El archivo mensual debe ser Excel (`.xlsx` o `.xlsm`) y contener una hoja llamada exactamente:

```text
Datos Generales
```

Columnas obligatorias:

```text
Patente
Pedimento
SeccionAduanera
TipoOperacion
ClaveDocumento
FechaPagoReal
```

## Ambiente virtual

Windows:

```powershell
py -3.12 -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

macOS/Linux para desarrollo:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Ejecutar

```bash
python src/main.py
```

## Generar exe en Windows

```powershell
pyinstaller --noconsole --onefile --name AuditorExpedientes src/main.py
```
