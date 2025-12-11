"""
Script para testear la extracción de metas del archivo Excel
"""
import sys
sys.path.insert(0, '.')
from app import read_xlsx_sheet_metas
import json

archivo_metas = r'C:\ws\sena\data\seguimiento_metas\sena-metas-app\data\aprendices\Seguimiento a Metas SENA 2025 V5 111125.xlsx'

print('=' * 100)
print('TESTING EXTRACCIÓN DE METAS')
print('=' * 100)

# Test 1: Extraer metas de CENTROS
print('\n1. Extrayendo metas de "5. FORMACIÓN X CTROS"...')
try:
    datos_centros = read_xlsx_sheet_metas(archivo_metas, "5. FORMACIÓN X CTROS")
    print(f'   [OK] Exito! {len(datos_centros)} registros extraidos')

    # Mostrar primer registro
    if datos_centros:
        print('\n   Primer registro de ejemplo:')
        primer_registro = datos_centros[0]
        print(f'   - PERIODO: {primer_registro.get("PERIODO")}')
        print(f'   - COD_REGIONAL: {primer_registro.get("COD_REGIONAL")}')
        print(f'   - REGIONAL: {primer_registro.get("REGIONAL")}')
        print(f'   - COD_CENTRO: {primer_registro.get("COD_CENTRO")}')
        print(f'   - CENTRO: {primer_registro.get("CENTRO")}')

        # Mostrar campos de metas (M_*)
        campos_metas = {k: v for k, v in primer_registro.items() if k.startswith('M_')}
        print(f'\n   Campos de metas encontrados: {len(campos_metas)}')
        print(f'   Primeros 5 campos:')
        for i, (campo, valor) in enumerate(list(campos_metas.items())[:5], 1):
            print(f'      {i}. {campo}: {valor}')

        # Guardar muestra en JSON
        output_file = 'muestra_metas_centros.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(datos_centros[:3], f, ensure_ascii=False, indent=2)
        print(f'\n   Primeros 3 registros guardados en: {output_file}')

except Exception as e:
    print(f'   [ERROR] Error: {str(e)}')
    import traceback
    traceback.print_exc()

# Test 2: Extraer metas de REGIONAL
print('\n2. Extrayendo metas de "4. FORMACION X REGIONAL"...')
try:
    datos_regional = read_xlsx_sheet_metas(archivo_metas, "4. FORMACIÓN X REGIONAL")
    print(f'   [OK] Exito! {len(datos_regional)} registros extraidos')

    # Mostrar primer registro
    if datos_regional:
        print('\n   Primer registro de ejemplo:')
        primer_registro = datos_regional[0]
        print(f'   - PERIODO: {primer_registro.get("PERIODO")}')
        print(f'   - COD_REGIONAL: {primer_registro.get("COD_REGIONAL")}')
        print(f'   - REGIONAL: {primer_registro.get("REGIONAL")}')

        # Mostrar campos de metas (M_*)
        campos_metas = {k: v for k, v in primer_registro.items() if k.startswith('M_')}
        print(f'\n   Campos de metas encontrados: {len(campos_metas)}')
        print(f'   Primeros 5 campos:')
        for i, (campo, valor) in enumerate(list(campos_metas.items())[:5], 1):
            print(f'      {i}. {campo}: {valor}')

        # Guardar muestra en JSON
        output_file = 'muestra_metas_regional.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(datos_regional[:3], f, ensure_ascii=False, indent=2)
        print(f'\n   Primeros 3 registros guardados en: {output_file}')

except Exception as e:
    print(f'   [ERROR] Error: {str(e)}')
    import traceback
    traceback.print_exc()

print('\n' + '=' * 100)
print('TESTING COMPLETADO')
print('=' * 100)
