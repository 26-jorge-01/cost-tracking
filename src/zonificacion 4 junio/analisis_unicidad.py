import pandas as pd

base_path = 'D:/ICBF/cost-tracking/data/insumos 4 junio/'
abastecimiento_df = pd.read_excel(
    base_path + 'zonificación_abastecimiento_servicios_primera_infancia25052026.xlsx',
    sheet_name='Integrales-Convenios',
    dtype={'Codigo Unidad Servicio UDS': str}
)
cuentame_df = pd.read_excel(
    base_path + 'ICBFCUEUnidadesServicio (8).xlsx',
    sheet_name='ICBFCUEUnidadesServicio',
    dtype={'Código UDS': str}
)

print('=== ABASTECIMIENTO ===')
print(f'Total rows: {len(abastecimiento_df)}')
print(f'Unique UDS codes: {abastecimiento_df["Codigo Unidad Servicio UDS"].nunique()}')

# Find UDS codes that appear multiple times
dup_codes = abastecimiento_df['Codigo Unidad Servicio UDS'].value_counts()
dup_codes = dup_codes[dup_codes > 1]
print(f'\nUDS codes duplicados en abastecimiento: {len(dup_codes)}')
print(f'Total filas de esos duplicados: {dup_codes.sum()}')

# Show examples of duplicated UDS
print('\n--- Ejemplos de UDS duplicados en abastecimiento ---')
for code in dup_codes.index[:3]:
    subset = abastecimiento_df[abastecimiento_df['Codigo Unidad Servicio UDS'] == code]
    cols = ['Codigo Unidad Servicio UDS', 'Unidad Servicio UDS', 'SERVICIO 2026',
            'Componente para la UDS', 'Desde', 'Hasta', 'Cupos', 'LITERAL DE CONTRATACION']
    print(f'\nUDS: {code}')
    print(subset[cols].to_string())

# Check what combination makes a row unique
print('\n\n--- Buscando la clave de unicidad ---')
# Check different column combinations
combos = [
    ['Codigo Unidad Servicio UDS'],
    ['Codigo Unidad Servicio UDS', 'SERVICIO 2026'],
    ['Codigo Unidad Servicio UDS', 'Componente para la UDS'],
    ['Codigo Unidad Servicio UDS', 'SERVICIO 2026', 'Componente para la UDS'],
    ['Codigo Unidad Servicio UDS', 'Desde'],
    ['Codigo Unidad Servicio UDS', 'Desde', 'Hasta'],
    ['Codigo Unidad Servicio UDS', 'SERVICIO 2026', 'Desde', 'Hasta'],
    ['Codigo Unidad Servicio UDS', 'SERVICIO 2026', 'Componente para la UDS', 'Desde', 'Hasta'],
    ['Codigo Unidad Servicio UDS', 'LITERAL DE CONTRATACION'],
    ['Codigo Unidad Servicio UDS', 'Componente para la UDS', 'LITERAL DE CONTRATACION'],
]
for combo in combos:
    unique_count = abastecimiento_df.drop_duplicates(subset=combo).shape[0]
    dup_count = len(abastecimiento_df) - unique_count
    print(f'  {combo}: {unique_count} unicos, {dup_count} duplicados')

print('\n\n=== CUENTAME ===')
print(f'Total rows: {len(cuentame_df)}')
print(f'Unique Código UDS: {cuentame_df["Código UDS"].nunique()}')

dup_codes_c = cuentame_df['Código UDS'].value_counts()
dup_codes_c = dup_codes_c[dup_codes_c > 1]
print(f'Códigos duplicados en cuentame: {len(dup_codes_c)}')
print(f'Filas de duplicados: {dup_codes_c.sum()}')

# Show examples of duplicated UDS in cuentame
print('\n--- Ejemplos de UDS duplicados en cuentame ---')
for code in dup_codes_c.index[:3]:
    subset = cuentame_df[cuentame_df['Código UDS'] == code]
    cols = ['Código UDS', 'Unidad De Servicio (UDS)', 'Nombre Servicio',
            'No. Contrato', 'Nombre Entidad Contratista', 'Estado de la UDS']
    cols = [c for c in cols if c in cuentame_df.columns]
    print(f'\nUDS: {code}')
    print(subset[cols].to_string())

# Check cuentame uniqueness combos
print('\n\n--- Buscando clave de unicidad en cuentame ---')
combos_c = [
    ['Código UDS'],
    ['Código UDS', 'Nombre Servicio'],
    ['Código UDS', 'No. Contrato'],
    ['Código UDS', 'Nombre Servicio', 'No. Contrato'],
]
for combo in combos_c:
    valid_cols = [c for c in combo if c in cuentame_df.columns]
    unique_count = cuentame_df.drop_duplicates(subset=valid_cols).shape[0]
    dup_count = len(cuentame_df) - unique_count
    print(f'  {valid_cols}: {unique_count} unicos, {dup_count} duplicados')
