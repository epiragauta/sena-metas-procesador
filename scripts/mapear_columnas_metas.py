"""
Script para mapear columnas de Cupos (metas) con sus categorías
"""
import openpyxl
import json

archivo_path = r'C:\ws\sena\data\seguimiento_metas\sena-metas-app\data\aprendices\Seguimiento a Metas SENA 2025 V5 111125.xlsx'
wb = openpyxl.load_workbook(archivo_path, data_only=True, read_only=True)

# Encontrar hoja
sheet_name = '5. FORMACIÓN X CTROS'
ws = wb[sheet_name]

# Leer las filas importantes
fila_codigos = list(ws.iter_rows(min_row=8, max_row=8, values_only=True))[0]  # Códigos
fila_categorias = list(ws.iter_rows(min_row=9, max_row=9, values_only=True))[0]  # Categorías
fila_encabezados = list(ws.iter_rows(min_row=10, max_row=10, values_only=True))[0]  # Cupos/Ejecución/%

# Primeras columnas de identificación (Regional, Centro, etc.)
print('=' * 100)
print('COLUMNAS DE IDENTIFICACIÓN (primeras 4 columnas):')
print('=' * 100)
for i in range(4):
    print(f'Col {openpyxl.utils.get_column_letter(i+1)}: ')
    print(f'  Fila 8: {fila_codigos[i]}')
    print(f'  Fila 9: {fila_categorias[i]}')
    print(f'  Fila 10: {fila_encabezados[i]}')
    print()

# Encontrar todas las columnas de "Cupos" y su mapeo
print('=' * 100)
print('MAPEO DE COLUMNAS "Cupos" (METAS):')
print('=' * 100)

mapeo_cupos = []
for i, encabezado in enumerate(fila_encabezados):
    if encabezado and 'Cupos' in str(encabezado):
        categoria = fila_categorias[i] if i < len(fila_categorias) else None
        codigo = fila_codigos[i] if i < len(fila_codigos) else None

        # Información de ejecución y porcentaje (siguiente y siguiente siguiente columna)
        ejecucion_header = fila_encabezados[i+1] if i+1 < len(fila_encabezados) else None
        porcentaje_header = fila_encabezados[i+2] if i+2 < len(fila_encabezados) else None

        mapeo_cupos.append({
            'columna': openpyxl.utils.get_column_letter(i+1),
            'indice': i + 1,
            'codigo': codigo,
            'categoria': categoria,
            'header_cupos': encabezado,
            'header_ejecucion': ejecucion_header,
            'header_porcentaje': porcentaje_header
        })

        print(f'\n{len(mapeo_cupos)}. Columna {openpyxl.utils.get_column_letter(i+1)} (índice {i+1}):')
        print(f'   Código: {codigo}')
        print(f'   Categoría: {categoria}')
        print(f'   Cupos (col {i+1}): {encabezado}')
        print(f'   Ejecución (col {i+2}): {ejecucion_header}')
        print(f'   % (col {i+3}): {porcentaje_header}')

print(f'\n\nTotal de columnas de metas encontradas: {len(mapeo_cupos)}')

# Guardar mapeo en JSON
output_file = 'mapeo_columnas_metas.json'
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump({
        'sheet_name': sheet_name,
        'header_row': 10,
        'categoria_row': 9,
        'codigo_row': 8,
        'data_start_row': 11,
        'columnas_identificacion': {
            'col_A': {'fila_8': fila_codigos[0], 'fila_9': fila_categorias[0], 'fila_10': fila_encabezados[0]},
            'col_B': {'fila_8': fila_codigos[1], 'fila_9': fila_categorias[1], 'fila_10': fila_encabezados[1]},
            'col_C': {'fila_8': fila_codigos[2], 'fila_9': fila_categorias[2], 'fila_10': fila_encabezados[2]},
            'col_D': {'fila_8': fila_codigos[3], 'fila_9': fila_categorias[3], 'fila_10': fila_encabezados[3]},
        },
        'columnas_metas': mapeo_cupos
    }, f, ensure_ascii=False, indent=2)

print(f'\n\nMapeo guardado en: {output_file}')

wb.close()
