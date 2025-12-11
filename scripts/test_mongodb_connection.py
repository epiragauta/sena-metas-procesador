"""
Script para probar la conexión a MongoDB y diagnosticar problemas.
"""
from pymongo import MongoClient
from dotenv import load_dotenv
import os

# Cargar variables de entorno
load_dotenv()

MONGODB_URI = os.getenv("MONGODB_URI")

print("=" * 80)
print("DIAGNÓSTICO DE CONEXIÓN A MONGODB")
print("=" * 80)

# Mostrar URI (ocultando contraseña)
uri_parts = MONGODB_URI.split(":")
if len(uri_parts) >= 3:
    masked_uri = uri_parts[0] + ":" + uri_parts[1] + ":<PASSWORD>@" + MONGODB_URI.split("@")[1]
else:
    masked_uri = "Error al parsear URI"

print(f"\n1. URI configurado (con contraseña oculta):")
print(f"   {masked_uri}")

print(f"\n2. Intentando conectar a MongoDB...")

try:
    # Intentar conexión básica
    client = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=5000)

    # Forzar conexión
    client.admin.command('ping')

    print("   ✓ Conexión exitosa!")

    # Listar bases de datos disponibles
    print(f"\n3. Bases de datos disponibles:")
    dbs = client.list_database_names()
    for db in dbs:
        print(f"   - {db}")

    # Intentar acceder a la base de datos sena_metas
    print(f"\n4. Intentando acceder a 'sena_metas':")
    db = client.get_database("sena_metas")
    collections = db.list_collection_names()
    print(f"   ✓ Acceso exitoso!")
    print(f"   Colecciones encontradas: {len(collections)}")
    for col in collections:
        print(f"   - {col}")

    client.close()
    print("\n" + "=" * 80)
    print("DIAGNÓSTICO COMPLETADO - TODO OK")
    print("=" * 80)

except Exception as e:
    print(f"   ✗ Error: {str(e)}")
    print(f"\n" + "=" * 80)
    print("POSIBLES CAUSAS DEL ERROR:")
    print("=" * 80)
    print("""
1. Usuario o contraseña incorrectos:
   - Verifica en MongoDB Atlas > Database Access
   - El usuario debe ser: Vercel-Admin-seguimiento-metas-mongodb

2. Usuario no tiene permisos:
   - En MongoDB Atlas > Database Access > tu usuario
   - Debe tener rol "readWrite" o "dbAdmin" para la base 'sena_metas'
   - O rol "Atlas admin" global

3. IP no está en la whitelist:
   - MongoDB Atlas > Network Access
   - Debe tener tu IP o 0.0.0.0/0 (permitir desde cualquier lugar)

4. Cluster pausado o no disponible:
   - Verifica en MongoDB Atlas que el cluster esté activo

5. Contraseña con caracteres especiales:
   - Si la contraseña tiene caracteres especiales, deben estar codificados

RECOMENDACIONES:
- Crea un nuevo usuario de base de datos en MongoDB Atlas
- Asigna rol "Atlas admin" o "readWrite" para la base 'sena_metas'
- Actualiza el URI en el archivo .env con las nuevas credenciales
""")
