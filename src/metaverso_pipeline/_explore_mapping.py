import pandas as pd, warnings
warnings.filterwarnings('ignore')

meta = pd.read_excel(r'data\insumos metaverso\Metaverso 2026.xlsx', sheet_name='ZonificacionPais')
cols_meta = list(meta.columns)

monetary = ['APALANCAMIENTO NUEVO DOS DIAS', 'VALOR TOTAL 2026 NUEVO', 'VALOR SOLO 2026',
            'APALANCA MAS VIGENCIA NUEVO 2026', 'Costo total hasta Julio ajuste operador alimentos']

def find_col(target, cols):
    t_norm = target.replace(' ', '').upper()
    for c in cols:
        c_norm = c.replace('\n', '').replace(' ', '').upper()
        if t_norm in c_norm or c_norm in t_norm:
            return c
    return None

print("=== CONCAT -> ZonificacionPais ===")
concat_to_meta = {
    'Regional_UDS': 'Regional UDS',
    'Centro_Zonal_UDS': 'Centro Zonal UDS',
    'Municipio_UDS': 'Municipio UDS',
    'ZONA': 'ZONA 2026',
    'Codigo_UDS': 'Codigo Unidad Servicio UDS',
    'UDS': 'Unidad Servicio UDS',
    'LITERAL_DE_CONTRATACION': 'TIPO DE CONTRATACION 2026',
    'Servicio': 'SERVICIO 2026',
    'Cupos': 'Cupos a Programar 2026',
    'Madres_Unds': 'Cantidad Madres Comunitarias en la UDS',
    'Componente_para_la_UDS': 'COMPONENTE',
    'Abrev': 'Modalidad 2026',
}
for c, m in concat_to_meta.items():
    found = find_col(m, cols_meta)
    print(f"  CONCAT.{c} -> META.{m} -> {found}")

print()
print("=== UDS_15052026 -> ZonificacionPais ===")
uds_to_meta = {
    'RegionalUDS': 'Regional UDS',
    'CentroZonalUDS': 'Centro Zonal UDS',
    'DepartamentoUDS': 'Departamento UDS',
    'MunicipioUDS': 'Municipio UDS',
    'CodigoUnidadServicioUDS': 'Codigo Unidad Servicio UDS',
    'UnidadServicioUDS': 'Unidad Servicio UDS',
    'BarrioUDS': 'Barrio UDS',
    'DireccionUDS': 'Direccion UDS',
    'EstadoUDS': 'Estado UDS',
    'CuposUDS': 'Cupos UDS ofertados 2025',
    'Servicio_2026': 'SERVICIO 2026',
    'ZonaUbicacionUDS': 'Zona Ubicacion UDS',
    'NombrePropiedadInfraestructura': 'Propiedad de la infraestructura',
    'EntidadContratista': 'CONTRATISTA 2026',
    'NumeroDocumentoEC': 'NIT CONTRATISTA 2026',
    'NumeroContrato': 'Numero Contrato',
    'TipoOrganizacionEC': 'Tipo de Organizacion',
}
for u, m in uds_to_meta.items():
    found = find_col(m, cols_meta)
    print(f"  UDS.{u} -> META.{m} -> {found}")

print()
print("=== COLUMNAS SIN MAPEAR ===")
cols_no_monetary = [c for c in cols_meta if c not in monetary]
mapped = set()
for m in list(concat_to_meta.values()) + list(uds_to_meta.values()):
    found = find_col(m, cols_meta)
    if found:
        mapped.add(found)
unmapped = [c for c in cols_no_monetary if c not in mapped]
print(f"Total no monetarias: {len(cols_no_monetary)}")
print(f"Mapeadas: {len(mapped)}")
print(f"Sin mapear: {len(unmapped)}")
for c in unmapped:
    print(f"  {repr(c)}")
