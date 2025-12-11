# API de Metas SENA - Documentación

## Descripción General

Esta API proporciona endpoints para cargar y consultar datos de metas del SENA desde archivos Excel (.xlsx). Los datos se almacenan en MongoDB y se exponen a través de endpoints REST.

## Base URL

**Desarrollo:** `http://localhost:8000`
**Producción:** `https://sena-metas-procesador.vercel.app`

---

## Endpoints

### 1. Cargar Metas desde Excel

**POST** `/upload-metas`

Carga un archivo Excel de metas SENA y procesa las hojas "4. FORMACIÓN X REGIONAL" y "5. FORMACIÓN X CTROS" para extraer datos de metas (columnas "Cupos").

#### Request

- **Content-Type:** `multipart/form-data`
- **Body:**
  - `file`: Archivo Excel (.xlsx) - Ejemplo: "Seguimiento a Metas SENA 2025 V5 111125.xlsx"

#### Response

```json
{
  "file_name": "Seguimiento a Metas SENA 2025 V5 111125.xlsx",
  "processing_date": "2025-11-22T13:49:27.744456",
  "collections_processed": 2,
  "errors": 0,
  "details": [
    {
      "sheet_name": "5. FORMACIÓN X CTROS",
      "collection_name": "metas_fpi_centros",
      "records_inserted": 120
    },
    {
      "sheet_name": "4. FORMACIÓN X REGIONAL",
      "collection_name": "metas_fpi_regional",
      "records_inserted": 38
    }
  ],
  "message": "Archivo procesado exitosamente. 2 colecciones actualizadas."
}
```

#### Ejemplo de uso (cURL)

```bash
curl -X POST "http://localhost:8000/upload-metas" \
  -F "file=@/ruta/a/Seguimiento a Metas SENA 2025 V5 111125.xlsx"
```

---

### 2. Consultar Metas de Centros

**GET** `/mongodb/collections/metas_fpi_centros`

Obtiene datos de metas por centro de formación.

#### Query Parameters

- `offset` (opcional): Número de registros a omitir (default: 0)
- `limit` (opcional): Número máximo de registros a retornar (default: todos)

#### Response

```json
{
  "collection_name": "metas_fpi_centros",
  "total_records": 120,
  "returned_records": 120,
  "offset": 0,
  "limit": null,
  "data": [
    {
      "PERIODO": "2025",
      "COD_REGIONAL": 5,
      "REGIONAL": "REGIONAL ANTIOQUIA",
      "COD_CENTRO": 9101,
      "CENTRO": "CENTRO DE LOS RECURSOS NATURALES RENOVABLES - LA SALADA",
      "M_TEC_REG_PRE": 1863.0,
      "M_TEC_REG_VIR": 622.0,
      "M_TEC_REG_A_D": 602.0,
      "M_TEC_CAMPESE": 62.0,
      "M_TEC_FULL_PO": 0.0,
      "M_TECNOLOGOS": 3149.0,
      "M_EDU_SUPERIO": 3149.0,
      "M_OPE_REGULAR": 150.0,
      "M_OPE_CAMPESE": 15.0,
      "M_OPE_FULL_PO": 15.0,
      "M_SUB_TOT_OPE": 180.0,
      "M_AUX_REGULAR": 48.0,
      "M_AUX_CAMPESE": 0.0,
      "M_AUX_FULL_PO": 0.0,
      "M_SUB_TOT_AUX": 48.0,
      "M_TCO_REG_PRE": 743.0,
      "M_TCO_REG_VIR": 189.0,
      "M_TCO_CAMPESE": 265.0,
      "M_TCO_FULL_PO": 15.0,
      "M_TCO_ART_MED": 2976.0,
      "M_SUB_TCO_LAB": 4188.0,
      "M_profundizacion_tecnica_t": 0.0,
      "M_TOT_FOR_LAB": 4416.0,
      "M_formacion_complementaria_virtual_sin_bilinguismo_g": 11440.0,
      "M_formacion_complementaria_presencial_sin_bilinguismo_h": 12350.0,
      "M_programa_de_bilinguismo_virtual_i": 3840.0,
      "M_programa_de_bilinguismo_presencial_j": 0.0,
      "M_subtotal_programa_de_bilinguismo_k_i_j": 3840.0,
      "M_COM_CAMPESE": 8733.0,
      "M_COM_FULL_PO": 100.0,
      "M_total_formacion_complementaria_n_g_h_k_l_m": 36463.0,
      "M_TOT_PROF_IN": 16091.0
    }
    // ... más registros
  ]
}
```

#### Ejemplo de uso (cURL)

```bash
# Obtener todos los registros
curl "http://localhost:8000/mongodb/collections/metas_fpi_centros"

# Obtener solo 10 registros
curl "http://localhost:8000/mongodb/collections/metas_fpi_centros?limit=10"

# Paginación: obtener registros del 20 al 30
curl "http://localhost:8000/mongodb/collections/metas_fpi_centros?offset=20&limit=10"
```

---

### 3. Consultar Metas Regionales

**GET** `/mongodb/collections/metas_fpi_regional`

Obtiene datos de metas consolidadas por regional.

#### Query Parameters

- `offset` (opcional): Número de registros a omitir (default: 0)
- `limit` (opcional): Número máximo de registros a retornar (default: todos)

#### Response

```json
{
  "collection_name": "metas_fpi_regional",
  "total_records": 38,
  "returned_records": 38,
  "offset": 0,
  "limit": null,
  "data": [
    {
      "PERIODO": "2025",
      "COD_REGIONAL": 5,
      "REGIONAL": "REGIONAL ANTIOQUIA",
      "COD_CENTRO": 44382,
      "CENTRO": 41751,
      "M_tecnologos_regular_presencial": 44382.0,
      "M_TEC_REG_VIR": 22934.0,
      "M_TEC_REG_A_D": 4983.0,
      "M_TEC_CAMPESE": 271.0,
      "M_TEC_FULL_PO": 0.0,
      "M_total_tecnologos_e": 72570.0,
      "M_total_educacion_superior_e": 72570.0,
      "M_OPE_REGULAR": 3653.0,
      "M_OPE_CAMPESE": 274.0,
      "M_OPE_FULL_PO": 101.0,
      "M_total_operarios_b": 4028.0,
      "M_AUX_REGULAR": 633.0,
      "M_AUX_CAMPESE": 136.0,
      "M_AUX_FULL_PO": 161.0,
      "M_total_auxiliares_a": 930.0,
      "M_TCO_REG_PRE": 32065.0,
      "M_TCO_REG_VIR": 9548.0,
      "M_TCO_CAMPESE": 3177.0,
      "M_TCO_FULL_PO": 165.0,
      "M_TCO_ART_MED": 49794.0,
      "M_total_tecnico_laboral_c": 94749.0,
      "M_total_profundizacion_tecnica_t": 30.0,
      "M_TOT_FOR_LAB": 99737.0,
      "M_formacion_complementaria_virtual_sin_bilinguismo_g": 235360.0,
      "M_formacion_complementaria_presencial_sin_bilinguismo_h": 260190.0,
      "M_programa_de_bilinguismo_virtual_i": 86160.0,
      "M_programa_de_bilinguismo_presencial_j": 12770.0,
      "M_total_programa_de_bilinguismo_k": 98930.0,
      "M_COM_CAMPESE": 82729.0,
      "M_COM_FULL_PO": 7040.0,
      "M_total_formacion_complementaria_n_g_h_k_l_m_incluye_los_cupos_de_formacion_continua_especial_campesina": 684249.0,
      "M_TOT_PROF_IN": 354002.0
    }
    // ... más registros
  ]
}
```

#### Ejemplo de uso (cURL)

```bash
# Obtener todos los registros
curl "http://localhost:8000/mongodb/collections/metas_fpi_regional"

# Obtener solo 5 registros
curl "http://localhost:8000/mongodb/collections/metas_fpi_regional?limit=5"
```

---

## Estructura de Datos

### Campos de Identificación (Centros)

- `PERIODO`: Año de las metas (e.g., "2025")
- `COD_REGIONAL`: Código numérico de la regional
- `REGIONAL`: Nombre de la regional
- `COD_CENTRO`: Código numérico del centro
- `CENTRO`: Nombre del centro de formación

### Campos de Identificación (Regional)

- `PERIODO`: Año de las metas (e.g., "2025")
- `COD_REGIONAL`: Código numérico de la regional
- `REGIONAL`: Nombre de la regional
- `COD_CENTRO`: Código consolidado (no aplicable para regional)
- `CENTRO`: Código consolidado (no aplicable para regional)

### Campos de Metas (prefijo M_)

Todos los campos de metas comienzan con el prefijo `M_` y representan cupos (metas) en diferentes categorías de formación:

#### Educación Superior (Tecnólogos)
- `M_TEC_REG_PRE`: Tecnólogos Regular - Presencial
- `M_TEC_REG_VIR`: Tecnólogos Regular - Virtual
- `M_TEC_REG_A_D`: Tecnólogos Regular - A Distancia
- `M_TEC_CAMPESE`: Tecnólogos CampeSENA
- `M_TEC_FULL_PO`: Tecnólogos Full Popular
- `M_TECNOLOGOS`: SubTotal Tecnólogos
- `M_EDU_SUPERIO`: Total Educación Superior

#### Formación Laboral - Operarios
- `M_OPE_REGULAR`: Operarios Regular
- `M_OPE_CAMPESE`: Operarios CampeSENA
- `M_OPE_FULL_PO`: Operarios Full Popular
- `M_SUB_TOT_OPE`: SubTotal Operarios

#### Formación Laboral - Auxiliares
- `M_AUX_REGULAR`: Auxiliares Regular
- `M_AUX_CAMPESE`: Auxiliares CampeSENA
- `M_AUX_FULL_PO`: Auxiliares Full Popular
- `M_SUB_TOT_AUX`: SubTotal Auxiliares

#### Formación Laboral - Técnicos
- `M_TCO_REG_PRE`: Técnico Laboral Regular - Presencial
- `M_TCO_REG_VIR`: Técnico Laboral Regular - Virtual
- `M_TCO_CAMPESE`: Técnico Laboral CampeSENA
- `M_TCO_FULL_PO`: Técnico Laboral Full Popular
- `M_TCO_ART_MED`: Técnico Laboral Articulación con la Media
- `M_SUB_TCO_LAB`: SubTotal Técnico Laboral

#### Totales Formación Laboral
- `M_TOT_FOR_LAB`: Total Formación Laboral

#### Formación Complementaria
- `M_COM_CAMPESE`: Formación Complementaria CampeSENA
- `M_COM_FULL_PO`: Formación Complementaria Full Popular
- `M_TOT_PROF_IN`: Total Formación Profesional Integral

---

## Colecciones MongoDB

### metas_fpi_centros
- **Descripción:** Metas de Formación Profesional Integral por Centro
- **Origen:** Hoja "5. FORMACIÓN X CTROS" del archivo Excel
- **Registros:** ~120 centros

### metas_fpi_regional
- **Descripción:** Metas de Formación Profesional Integral consolidadas por Regional
- **Origen:** Hoja "4. FORMACIÓN X REGIONAL" del archivo Excel
- **Registros:** 38 regionales

---

## Proceso de Actualización

1. **Recepción del archivo:** Se recibe el archivo Excel actualizado (2 veces por semana)
2. **Validación:** El sistema valida que sea un archivo .xlsx válido
3. **Extracción:** Se extraen los datos de las hojas:
   - "4. FORMACIÓN X REGIONAL"
   - "5. FORMACIÓN X CTROS"
4. **Normalización:** Los nombres de categorías se normalizan a campos M_*
5. **Almacenamiento:** Los datos se guardan en MongoDB (reemplazando datos anteriores)
6. **Disponibilidad:** Los datos quedan disponibles inmediatamente via API

---

## Códigos de Estado HTTP

- `200 OK` - Solicitud exitosa
- `400 Bad Request` - Error en los parámetros o archivo inválido
- `404 Not Found` - Colección no encontrada
- `500 Internal Server Error` - Error del servidor

---

## Notas Importantes

1. **Reemplazo de datos:** Al cargar un nuevo archivo, los datos anteriores se eliminan completamente
2. **Periodo fijo:** Actualmente el periodo está hardcoded a "2025"
3. **Normalización de nombres:** Los nombres de categorías del Excel se normalizan a snake_case con prefijo M_
4. **Campos vacíos:** Los valores vacíos en Excel se convierten en `null` o `0.0` según el contexto
5. **Encoding:** Todos los textos están en UTF-8

---

## Próximos Pasos

- **Fase 2:** Crear servicios Angular para consumir esta API
- **Fase 3:** Implementar transformación de datos en el frontend
- **Fase 4:** Integración híbrida (API + fallback a JSON)
- **Fase 5:** Mapeo de tablas adicionales
- **Fase 6:** Despliegue y monitoreo

---

## Contacto y Soporte

Para más información sobre esta API, consulta el repositorio del proyecto o el archivo `ESTRATEGIA_IMPLEMENTACION_API.md`.
