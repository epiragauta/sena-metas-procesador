"""
Script para testear la extracción de metas del archivo Excel.
Valida que la estructura de datos sea correcta para ambas hojas (CENTROS y REGIONAL).
"""
import sys
sys.path.insert(0, '.')
from app import read_xlsx_sheet_metas
import json

archivo_metas = r'C:\ws\sena\data\seguimiento_metas\sena-metas-app\data\aprendices\Seguimiento a Metas SENA 2025 V5 -25112025.xlsx'

print('=' * 100)
print('TESTING EXTRACCIÓN DE METAS')
print('=' * 100)

# Test 1: Extraer metas de CENTROS
print('\n1. Extrayendo metas de "5. FORMACIÓN X CTROS"...')
try:
    datos_centros = read_xlsx_sheet_metas(archivo_metas, "5. FORMACIÓN X CTROS")
    print(f'   [OK] Exito! {len(datos_centros)} registros extraidos')

    # Validaciones
    validaciones_ok = True

    if datos_centros:
        print('\n   Primer registro de ejemplo:')
        primer_registro = datos_centros[0]
        print(f'   - PERIODO: {primer_registro.get("PERIODO")}')
        print(f'   - COD_REGIONAL: {primer_registro.get("COD_REGIONAL")}')
        print(f'   - REGIONAL: {primer_registro.get("REGIONAL")}')
        print(f'   - COD_CENTRO: {primer_registro.get("COD_CENTRO")}')
        print(f'   - CENTRO: {primer_registro.get("CENTRO")}')

        # VALIDACIÓN 1: COD_CENTRO y CENTRO deben tener valores (no null)
        print('\n   VALIDACIONES:')
        if primer_registro.get("COD_CENTRO") is not None and primer_registro.get("CENTRO") is not None:
            print(f'   ✓ COD_CENTRO y CENTRO tienen valores (correcto para hoja CTROS)')
        else:
            print(f'   ✗ ERROR: COD_CENTRO o CENTRO son null (incorrecto para hoja CTROS)')
            validaciones_ok = False

        # Mostrar campos de metas (M_*)
        campos_metas = {k: v for k, v in primer_registro.items() if k.startswith('M_')}
        print(f'\n   Campos de metas encontrados: {len(campos_metas)}')

        # VALIDACIÓN 2: Verificar que haya campos mapeados correctamente
        campos_esperados = ['M_TEC_REG_PRE', 'M_TEC_REG_VIR', 'M_TEC_REG_A_D']
        campos_encontrados = [c for c in campos_esperados if c in campos_metas]
        if len(campos_encontrados) == len(campos_esperados):
            print(f'   ✓ Campos mapeados correctamente: {", ".join(campos_encontrados)}')
        else:
            print(f'   ✗ ERROR: Faltan campos mapeados. Encontrados: {campos_encontrados}')
            validaciones_ok = False

        print(f'\n   Primeros 5 campos:')
        for i, (campo, valor) in enumerate(list(campos_metas.items())[:5], 1):
            print(f'      {i}. {campo}: {valor}')

        # Guardar muestra en JSON
        output_file = 'muestra_metas_centros.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(datos_centros[:3], f, ensure_ascii=False, indent=2)
        print(f'\n   Primeros 3 registros guardados en: {output_file}')

        if validaciones_ok:
            print('\n   ✅ TODAS LAS VALIDACIONES PASARON')
        else:
            print('\n   ❌ ALGUNAS VALIDACIONES FALLARON')

except Exception as e:
    print(f'   [ERROR] Error: {str(e)}')
    import traceback
    traceback.print_exc()

# Test 2: Extraer metas de REGIONAL
print('\n2. Extrayendo metas de "4. FORMACIÓN X REGIONAL"...')
try:
    datos_regional = read_xlsx_sheet_metas(archivo_metas, "4. FORMACIÓN X REGIONAL")
    print(f'   [OK] Exito! {len(datos_regional)} registros extraidos')

    # Validaciones
    validaciones_ok = True

    if datos_regional:
        print('\n   Primer registro de ejemplo:')
        primer_registro = datos_regional[0]
        print(f'   - PERIODO: {primer_registro.get("PERIODO")}')
        print(f'   - COD_REGIONAL: {primer_registro.get("COD_REGIONAL")}')
        print(f'   - REGIONAL: {primer_registro.get("REGIONAL")}')
        print(f'   - COD_CENTRO: {primer_registro.get("COD_CENTRO")}')
        print(f'   - CENTRO: {primer_registro.get("CENTRO")}')

        # VALIDACIÓN 1: COD_CENTRO y CENTRO deben ser null para REGIONAL
        print('\n   VALIDACIONES:')
        if primer_registro.get("COD_CENTRO") is None and primer_registro.get("CENTRO") is None:
            print(f'   ✓ COD_CENTRO y CENTRO son null (correcto para hoja REGIONAL)')
        else:
            print(f'   ✗ ERROR: COD_CENTRO o CENTRO no son null (incorrecto para hoja REGIONAL)')
            print(f'      COD_CENTRO: {primer_registro.get("COD_CENTRO")}')
            print(f'      CENTRO: {primer_registro.get("CENTRO")}')
            validaciones_ok = False

        # Mostrar campos de metas (M_*)
        campos_metas = {k: v for k, v in primer_registro.items() if k.startswith('M_')}
        print(f'\n   Campos de metas encontrados: {len(campos_metas)}')

        # VALIDACIÓN 2: Verificar que los campos estén mapeados correctamente (no genéricos)
        campos_esperados = ['M_TEC_REG_PRE', 'M_TEC_REG_VIR', 'M_TEC_REG_A_D']
        campos_encontrados = [c for c in campos_esperados if c in campos_metas]

        if len(campos_encontrados) == len(campos_esperados):
            print(f'   ✓ Campos mapeados correctamente (con normalización de tildes): {", ".join(campos_encontrados)}')
        else:
            print(f'   ✗ ERROR: Faltan campos mapeados. Encontrados: {campos_encontrados}')
            validaciones_ok = False

        # VALIDACIÓN 3: Verificar que no haya campos genéricos (m_tecnologos_regular_presencial)
        campos_genericos = [k for k in campos_metas.keys() if 'tecnologos_regular_presencial' in k.lower()]
        if not campos_genericos:
            print(f'   ✓ No hay campos genéricos (normalización funcionó correctamente)')
        else:
            print(f'   ✗ ERROR: Se encontraron campos genéricos: {campos_genericos}')
            validaciones_ok = False

        # VALIDACIÓN 4: Verificar que M_TEC_REG_PRE tenga un valor razonable
        m_tec_reg_pre = campos_metas.get('M_TEC_REG_PRE', 0)
        if m_tec_reg_pre > 0:
            print(f'   ✓ M_TEC_REG_PRE tiene valor: {m_tec_reg_pre}')
        else:
            print(f'   ✗ ERROR: M_TEC_REG_PRE no tiene valor o es 0')
            validaciones_ok = False

        print(f'\n   Primeros 5 campos:')
        for i, (campo, valor) in enumerate(list(campos_metas.items())[:5], 1):
            print(f'      {i}. {campo}: {valor}')

        # Guardar muestra en JSON
        output_file = 'muestra_metas_regional.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(datos_regional[:3], f, ensure_ascii=False, indent=2)
        print(f'\n   Primeros 3 registros guardados en: {output_file}')

        if validaciones_ok:
            print('\n   ✅ TODAS LAS VALIDACIONES PASARON')
        else:
            print('\n   ❌ ALGUNAS VALIDACIONES FALLARON')

except Exception as e:
    print(f'   [ERROR] Error: {str(e)}')
    import traceback
    traceback.print_exc()

print('\n' + '=' * 100)
print('TESTING COMPLETADO')
print('=' * 100)
