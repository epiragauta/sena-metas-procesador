"""
API para procesar archivos XLSB y generar datos JSON.

Esta API permite:
- Subir archivos XLSB
- Listar las hojas disponibles (excluyendo hojas SQL)
- Obtener datos de cada hoja en formato JSON
- Exportar hojas a archivos JSON

Uso:
    uvicorn xlsb_api:app --reload --port 8000
"""

from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import Optional, List
from pyxlsb import open_workbook
from pymongo import MongoClient
from dotenv import load_dotenv
import json
import os
import tempfile
import shutil
import re
import unicodedata
from pathlib import Path
from datetime import datetime

# Cargar variables de entorno
load_dotenv()

# Conexión a MongoDB
MONGODB_URI = os.getenv("MONGODB_URI")
mongo_client = None
db = None

def get_database():
    """Obtiene la conexión a la base de datos MongoDB."""
    global mongo_client, db
    if mongo_client is None:
        mongo_client = MongoClient(MONGODB_URI)
        db = mongo_client.get_database("sena_metas")
    return db

def normalize_collection_name(sheet_name: str) -> str:
    """
    Normaliza el nombre de la hoja para usarlo como nombre de colección.
    Convierte a minúsculas y elimina caracteres especiales.
    """
    # Normalizar caracteres unicode (remover acentos)
    normalized = unicodedata.normalize('NFKD', sheet_name)
    normalized = normalized.encode('ASCII', 'ignore').decode('ASCII')
    # Convertir a minúsculas
    normalized = normalized.lower()
    # Reemplazar espacios y caracteres especiales por guiones bajos
    normalized = re.sub(r'[^a-z0-9]', '_', normalized)
    # Eliminar guiones bajos múltiples
    normalized = re.sub(r'_+', '_', normalized)
    # Eliminar guiones bajos al inicio y final
    normalized = normalized.strip('_')
    return normalized

def is_ejecucion_fpi_file(filename: str) -> bool:
    """Verifica si el archivo es de tipo 'ejecución FPI'."""
    normalized = filename.lower()
    return 'ejecucion fpi' in normalized or 'ejecución fpi' in normalized

app = FastAPI(
    title="XLSB to JSON API",
    description="API para convertir archivos XLSB a JSON",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json"
)

# Configurar CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Directorio temporal para archivos subidos y generados
UPLOAD_DIR = Path(tempfile.gettempdir()) / "xlsb_api_uploads"
OUTPUT_DIR = Path(tempfile.gettempdir()) / "xlsb_api_output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Almacén de archivos subidos (en memoria para demo)
uploaded_files = {}


def read_xlsb_sheet(file_path: str, sheet_name: str) -> List[dict]:
    """
    Lee una hoja de un archivo XLSB y retorna los datos como lista de diccionarios.

    Args:
        file_path: Ruta al archivo XLSB
        sheet_name: Nombre de la hoja a leer

    Returns:
        Lista de diccionarios con los datos de la hoja
    """
    data = []
    with open_workbook(file_path) as wb:
        with wb.get_sheet(sheet_name) as sheet:
            rows = list(sheet.rows())
            if not rows:
                return data

            # Primera fila como encabezados
            headers = []
            for cell in rows[0]:
                value = cell.v
                if value is None:
                    value = f"column_{len(headers)}"
                headers.append(str(value).strip())

            # Procesar filas de datos
            for row in rows[1:]:
                row_data = {}
                for i, cell in enumerate(row):
                    if i < len(headers):
                        value = cell.v
                        # Convertir valores numéricos
                        if isinstance(value, (int, float)):
                            row_data[headers[i]] = value
                        elif value is not None:
                            row_data[headers[i]] = str(value).strip()
                        else:
                            row_data[headers[i]] = None

                # Solo agregar filas que tengan al menos un valor no nulo
                if any(v is not None for v in row_data.values()):
                    data.append(row_data)

    return data


def get_sheet_names(file_path: str, exclude_sql: bool = True) -> List[str]:
    """
    Obtiene los nombres de las hojas de un archivo XLSB.

    Args:
        file_path: Ruta al archivo XLSB
        exclude_sql: Si es True, excluye hojas con "SQL" en el nombre

    Returns:
        Lista de nombres de hojas
    """
    with open_workbook(file_path) as wb:
        sheets = wb.sheets
        if exclude_sql:
            sheets = [s for s in sheets if 'SQL' not in s.upper()]
        return sheets


@app.get("/")
async def root():
    """Endpoint raíz con información de la API."""
    return {
        "message": "XLSB to JSON API",
        "version": "1.0.0",
        "endpoints": {
            "POST /upload": "Subir archivo XLSB",
            "GET /files": "Listar archivos subidos",
            "GET /files/{file_id}/sheets": "Listar hojas de un archivo",
            "GET /files/{file_id}/sheets/{sheet_name}": "Obtener datos de una hoja",
            "POST /files/{file_id}/export": "Exportar todas las hojas a JSON",
            "GET /download/{filename}": "Descargar archivo generado"
        }
    }


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """
    Sube un archivo XLSB para procesamiento.
    Si el nombre contiene 'ejecución FPI', guarda los datos en MongoDB.

    Returns:
        ID del archivo y lista de hojas disponibles
    """
    if not file.filename.endswith('.xlsb'):
        raise HTTPException(status_code=400, detail="Solo se permiten archivos .xlsb")

    # Generar ID único
    file_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}"
    file_path = UPLOAD_DIR / file_id

    # Guardar archivo
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Obtener hojas
    try:
        sheets = get_sheet_names(str(file_path))
    except Exception as e:
        os.remove(file_path)
        raise HTTPException(status_code=400, detail=f"Error al leer archivo: {str(e)}")

    # Registrar archivo
    uploaded_files[file_id] = {
        "path": str(file_path),
        "original_name": file.filename,
        "sheets": sheets,
        "uploaded_at": datetime.now().isoformat()
    }

    # Si es archivo de ejecución FPI, guardar en MongoDB
    mongodb_collections = []
    if is_ejecucion_fpi_file(file.filename):
        try:
            database = get_database()
            for sheet_name in sheets:
                # Leer datos de la hoja
                data = read_xlsb_sheet(str(file_path), sheet_name)

                if data:
                    # Crear nombre de colección normalizado
                    collection_name = f"ejecucion_fpi_{normalize_collection_name(sheet_name)}"

                    # Obtener colección y limpiar datos anteriores
                    collection = database[collection_name]
                    collection.delete_many({})

                    # Insertar nuevos datos
                    collection.insert_many(data)

                    mongodb_collections.append({
                        "sheet_name": sheet_name,
                        "collection_name": collection_name,
                        "records_inserted": len(data)
                    })
        except Exception as e:
            # No fallar el upload si MongoDB tiene problemas
            mongodb_collections.append({
                "error": f"Error al guardar en MongoDB: {str(e)}"
            })

    response = {
        "file_id": file_id,
        "original_name": file.filename,
        "sheets": sheets,
        "message": f"Archivo subido exitosamente. {len(sheets)} hojas disponibles (excluyendo SQL)."
    }

    if mongodb_collections:
        response["mongodb_collections"] = mongodb_collections
        response["message"] += f" Datos guardados en {len([c for c in mongodb_collections if 'collection_name' in c])} colecciones de MongoDB."

    return response


@app.get("/files")
async def list_files():
    """Lista todos los archivos subidos."""
    return {
        "files": [
            {
                "file_id": fid,
                "original_name": info["original_name"],
                "sheets": info["sheets"],
                "uploaded_at": info["uploaded_at"]
            }
            for fid, info in uploaded_files.items()
        ]
    }


@app.get("/files/{file_id}/sheets")
async def get_sheets(file_id: str):
    """
    Lista las hojas disponibles en un archivo.

    Args:
        file_id: ID del archivo subido
    """
    if file_id not in uploaded_files:
        raise HTTPException(status_code=404, detail="Archivo no encontrado")

    return {
        "file_id": file_id,
        "sheets": uploaded_files[file_id]["sheets"]
    }


@app.get("/files/{file_id}/sheets/{sheet_name}")
async def get_sheet_data(
    file_id: str,
    sheet_name: str,
    limit: Optional[int] = Query(None, description="Límite de registros"),
    offset: Optional[int] = Query(0, description="Offset de registros")
):
    """
    Obtiene los datos de una hoja en formato JSON.

    Args:
        file_id: ID del archivo subido
        sheet_name: Nombre de la hoja
        limit: Número máximo de registros a retornar
        offset: Número de registros a saltar
    """
    if file_id not in uploaded_files:
        raise HTTPException(status_code=404, detail="Archivo no encontrado")

    if sheet_name not in uploaded_files[file_id]["sheets"]:
        raise HTTPException(status_code=404, detail=f"Hoja '{sheet_name}' no encontrada")

    try:
        data = read_xlsb_sheet(uploaded_files[file_id]["path"], sheet_name)

        # Aplicar paginación
        total = len(data)
        if offset:
            data = data[offset:]
        if limit:
            data = data[:limit]

        return {
            "file_id": file_id,
            "sheet_name": sheet_name,
            "total_records": total,
            "returned_records": len(data),
            "offset": offset,
            "limit": limit,
            "data": data
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al leer hoja: {str(e)}")


@app.post("/files/{file_id}/export")
async def export_all_sheets(file_id: str):
    """
    Exporta todas las hojas a archivos JSON individuales.

    Args:
        file_id: ID del archivo subido

    Returns:
        Lista de archivos generados con sus URLs de descarga
    """
    if file_id not in uploaded_files:
        raise HTTPException(status_code=404, detail="Archivo no encontrado")

    exported_files = []
    file_info = uploaded_files[file_id]
    base_name = Path(file_info["original_name"]).stem

    for sheet_name in file_info["sheets"]:
        try:
            data = read_xlsb_sheet(file_info["path"], sheet_name)

            output = {
                "sheet_name": sheet_name,
                "total_records": len(data),
                "generated_at": datetime.now().isoformat(),
                "data": data
            }

            # Guardar archivo
            filename = f"{base_name}_{sheet_name}.json"
            output_path = OUTPUT_DIR / filename

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(output, f, ensure_ascii=False, indent=2)

            exported_files.append({
                "sheet_name": sheet_name,
                "filename": filename,
                "records": len(data),
                "download_url": f"/download/{filename}"
            })

        except Exception as e:
            exported_files.append({
                "sheet_name": sheet_name,
                "error": str(e)
            })

    return {
        "file_id": file_id,
        "exported_files": exported_files
    }


@app.get("/download/{filename}")
async def download_file(filename: str):
    """
    Descarga un archivo generado.

    Args:
        filename: Nombre del archivo a descargar
    """
    file_path = OUTPUT_DIR / filename

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado")

    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/json"
    )


@app.delete("/files/{file_id}")
async def delete_file(file_id: str):
    """
    Elimina un archivo subido.

    Args:
        file_id: ID del archivo a eliminar
    """
    if file_id not in uploaded_files:
        raise HTTPException(status_code=404, detail="Archivo no encontrado")

    # Eliminar archivo físico
    try:
        os.remove(uploaded_files[file_id]["path"])
    except:
        pass

    # Eliminar registro
    del uploaded_files[file_id]

    return {"message": f"Archivo {file_id} eliminado"}


# Endpoint adicional para procesar archivo local (útil para desarrollo)
@app.post("/process-local")
async def process_local_file(
    file_path: str = Query(..., description="Ruta al archivo XLSB local"),
    output_dir: Optional[str] = Query(None, description="Directorio de salida")
):
    """
    Procesa un archivo XLSB local y genera archivos JSON de salida.

    Args:
        file_path: Ruta absoluta al archivo XLSB
        output_dir: Directorio donde guardar los archivos generados
    """
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail=f"Archivo no encontrado: {file_path}")

    if not file_path.endswith('.xlsb'):
        raise HTTPException(status_code=400, detail="Solo se permiten archivos .xlsb")

    # Directorio de salida
    if output_dir:
        out_path = Path(output_dir)
    else:
        out_path = Path(file_path).parent / "output"

    out_path.mkdir(exist_ok=True)

    # Obtener hojas
    sheets = get_sheet_names(file_path)
    base_name = Path(file_path).stem

    exported_files = []

    for sheet_name in sheets:
        try:
            data = read_xlsb_sheet(file_path, sheet_name)

            output = {
                "sheet_name": sheet_name,
                "total_records": len(data),
                "generated_at": datetime.now().isoformat(),
                "data": data
            }

            # Guardar archivo
            filename = f"{base_name}_{sheet_name}.json"
            output_file = out_path / filename

            with open(output_file, "w", encoding="utf-8") as f:
                json.dump(output, f, ensure_ascii=False, indent=2)

            exported_files.append({
                "sheet_name": sheet_name,
                "filename": filename,
                "output_path": str(output_file),
                "records": len(data)
            })

        except Exception as e:
            exported_files.append({
                "sheet_name": sheet_name,
                "error": str(e)
            })

    return {
        "source_file": file_path,
        "output_directory": str(out_path),
        "sheets_processed": len(sheets),
        "exported_files": exported_files
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
