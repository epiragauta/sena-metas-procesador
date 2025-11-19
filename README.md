# API XLSB to JSON

API en Python para convertir archivos Excel binarios (.xlsb) a formato JSON.

## Instalación

```bash
cd api
pip install -r requirements.txt
```

## Uso

### Iniciar el servidor

```bash
uvicorn xlsb_api:app --reload --port 8000
```

O directamente:

```bash
python xlsb_api.py
```

La API estará disponible en `http://localhost:8000`

### Documentación interactiva

- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## Endpoints

### 1. Subir archivo XLSB

```bash
POST /upload
```

```bash
curl -X POST "http://localhost:8000/upload" \
  -F "file=@archivo.xlsb"
```

Respuesta:
```json
{
  "file_id": "20251118_143000_archivo.xlsb",
  "original_name": "archivo.xlsb",
  "sheets": ["REGIONAL", "CENTROS", "TP_REGIONAL"],
  "message": "Archivo subido exitosamente. 3 hojas disponibles (excluyendo SQL)."
}
```

### 2. Listar archivos subidos

```bash
GET /files
```

### 3. Listar hojas de un archivo

```bash
GET /files/{file_id}/sheets
```

### 4. Obtener datos de una hoja (JSON)

```bash
GET /files/{file_id}/sheets/{sheet_name}
```

Parámetros opcionales:
- `limit`: Número máximo de registros
- `offset`: Registros a saltar

Ejemplo:
```bash
curl "http://localhost:8000/files/{file_id}/sheets/REGIONAL?limit=10&offset=0"
```

Respuesta:
```json
{
  "file_id": "...",
  "sheet_name": "REGIONAL",
  "total_records": 33,
  "returned_records": 10,
  "offset": 0,
  "limit": 10,
  "data": [
    {
      "REGIONAL": "ANTIOQUIA",
      "CODIGO": 5,
      "META": 10000,
      "EJECUCION": 9500
    }
  ]
}
```

### 5. Exportar todas las hojas a archivos JSON

```bash
POST /files/{file_id}/export
```

Ejemplo:
```bash
curl -X POST "http://localhost:8000/files/{file_id}/export"
```

Respuesta:
```json
{
  "file_id": "...",
  "exported_files": [
    {
      "sheet_name": "REGIONAL",
      "filename": "archivo_REGIONAL.json",
      "records": 33,
      "download_url": "/download/archivo_REGIONAL.json"
    }
  ]
}
```

### 6. Descargar archivo generado

```bash
GET /download/{filename}
```

### 7. Procesar archivo local (desarrollo)

```bash
POST /process-local
```

Parámetros:
- `file_path`: Ruta absoluta al archivo XLSB
- `output_dir`: Directorio de salida (opcional)

Ejemplo:
```bash
curl -X POST "http://localhost:8000/process-local?file_path=C:/data/archivo.xlsb"
```

### 8. Eliminar archivo

```bash
DELETE /files/{file_id}
```

## Ejemplo completo

```python
import requests

# 1. Subir archivo
with open("plantilla.xlsb", "rb") as f:
    response = requests.post(
        "http://localhost:8000/upload",
        files={"file": f}
    )
    file_id = response.json()["file_id"]

# 2. Ver hojas disponibles
sheets = requests.get(f"http://localhost:8000/files/{file_id}/sheets").json()
print("Hojas:", sheets["sheets"])

# 3. Obtener datos de una hoja
data = requests.get(f"http://localhost:8000/files/{file_id}/sheets/REGIONAL").json()
print(f"Registros: {data['total_records']}")

# 4. Exportar todas las hojas a JSON
export = requests.post(f"http://localhost:8000/files/{file_id}/export").json()

# 5. Descargar archivos generados
for file_info in export["exported_files"]:
    if "download_url" in file_info:
        content = requests.get(f"http://localhost:8000{file_info['download_url']}").json()
        print(f"{file_info['sheet_name']}: {content['total_records']} registros")
```

## Estructura del JSON generado

```json
{
  "sheet_name": "REGIONAL",
  "total_records": 33,
  "generated_at": "2025-11-18T14:30:00",
  "data": [
    {
      "REGIONAL": "ANTIOQUIA",
      "CODIGO": 5,
      "META": 10000,
      "EJECUCION": 9500,
      "PORCENTAJE": 95.0
    }
  ]
}
```

## Notas

- Las hojas con "SQL" en el nombre son excluidas automáticamente
- Los archivos se almacenan temporalmente y deben eliminarse después de su uso
- Soporta paginación con `limit` y `offset` para hojas grandes
