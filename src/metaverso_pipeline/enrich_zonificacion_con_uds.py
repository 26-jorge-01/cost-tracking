import pandas as pd
import numpy as np
import os
import re
import unicodedata
import warnings
warnings.filterwarnings('ignore')

DIR_BASE = r"D:\ICBF\cost-tracking\data\insumos metaverso"
INPUT_CONCAT = os.path.join(DIR_BASE, "CONCAT_ZONIFICACION_METAVERSO_2026.xlsx")
INPUT_UDS = os.path.join(DIR_BASE, "UDS_31122025.xlsx")
INPUT_TEMPLATE = os.path.join(DIR_BASE, "Metaverso 2026.xlsx")
OUTPUT_FILE = os.path.join(DIR_BASE, "ZONIFICACION_ENRIQUECIDA_2026.xlsx")

SHEET_UDS = "UDS_31122026"
SHEET_TEMPLATE = "ZonificacionPais"

def log(msg):
    print(f"[LOG] {msg}")

def norm_key(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    s = re.sub(r'\.0$', '', s)
    return s if s else None

def clean_nit(x):
    if pd.isna(x):
        return None
    s = str(x).replace(',', '.').split('.')[0]
    s = ''.join(c for c in s if c.isdigit())
    return s if s else None

def normalize(s):
    return unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('ASCII').lower().replace(' ', '').replace('\n', '')

# ============================================================
log("PASO 1: Cargar template ZonificacionPais")
# ============================================================
meta = pd.read_excel(INPUT_TEMPLATE, sheet_name=SHEET_TEMPLATE)
TEMPLATE_COLS = list(meta.columns)
log(f"  {len(TEMPLATE_COLS)} columnas en template")

# Build normalized lookup for fuzzy matching
template_norm = {normalize(c): c for c in TEMPLATE_COLS}

# ============================================================
log("PASO 2: Cargar CONCAT")
# ============================================================
concat = pd.read_excel(INPUT_CONCAT)
log(f"  {len(concat)} filas, {len(concat.columns)} columnas")

# Quality fix 1: deduplicate true duplicates
concat["_k"] = concat["Codigo_UDS"].astype(str).str.strip()
concat["_k"] = concat["_k"].apply(lambda x: re.sub(r'\.0$', '', x) if x else x)

# Find true duplicates (same code, same Servicio, same Tipo_Modalidad)
concat["_dup_key"] = concat["_k"] + "|" + concat["Servicio"].fillna("").astype(str) + "|" + concat["Tipo_Modalidad"].fillna("").astype(str)
true_dups_before = concat["_dup_key"].duplicated(keep="first").sum()
concat = concat.drop_duplicates(subset="_dup_key", keep="first")
log(f"  True duplicates eliminados: {true_dups_before}")
log(f"  Filas despues de dedup: {len(concat)}")

# Flag "No aplica" codes
concat["_no_match"] = concat["_k"].apply(lambda x: True if x == "No aplica" else False)
no_aplica = concat["_no_match"].sum()
log(f"  Codigos 'No aplica': {no_aplica}")

# ============================================================
log("PASO 3: Cargar UDS_31122025")
# ============================================================
uds = pd.read_excel(INPUT_UDS, sheet_name=SHEET_UDS)
log(f"  {len(uds)} filas, {len(uds.columns)} columnas")

uds["_k"] = uds["CodigoUnidadServicioUDS"].astype(str).str.strip()
uds["_k"] = uds["_k"].apply(lambda x: re.sub(r'\.0$', '', x) if x else None)

# No duplicates in UDS, but ensure unique
uds = uds.drop_duplicates(subset="_k", keep="first")
log(f"  UDS keys unicas: {len(uds)}")

# ============================================================
log("PASO 4: Merge CONCAT <- UDS_31122025")
# ============================================================
merged = concat.merge(uds, on="_k", how="left", suffixes=("", "_uds"))
log(f"  Merge completado: {len(merged)} filas")

matched = merged["EntidadContratista"].notna().sum()
log(f"  Enlazados con UDS: {matched}/{len(merged)}")

# ============================================================
log("PASO 5: Construir mapping CONCAT -> ZonificacionPais")
# ============================================================
# Priority: CONCAT data takes precedence over UDS for overlapping fields
# We'll fill the template column by column

out = pd.DataFrame(index=range(len(merged)), columns=TEMPLATE_COLS)
out_filled = set()

# ---- MAP CONCAT FIELDS ----
concat_mapping = {
    "Regional_UDS": "Regional UDS",
    "Centro_Zonal_UDS": "Centro Zonal UDS",
    "Municipio_UDS": "Municipio UDS",
    "ZONA": "ZONA 2026",
    "Codigo_UDS": "Codigo Unidad Servicio UDS",
    "UDS": "Unidad Servicio UDS",
    "LITERAL_DE_CONTRATACION": "TIPO DE\nCONTRATACIÓN 2026",
    "Servicio": "SERVICIO\n2026",
    "Cupos": "Cupos a Programar 2026",
    "Madres_Unds": "Cantidad Madres\nComunitarias en la UDS",
    "Componente_para_la_UDS": "COMPONENTE",
    "Atenciones": None,
    "Atenciones_30042026": None,
    "Abrev": None,
    "LITERAL_DE_CONTRATACION_ALTERNATIVO": None,
}

log("  Mapeando campos desde CONCAT...")
for src_col, target_col in concat_mapping.items():
    if src_col not in merged.columns:
        continue
    if target_col is None:
        continue
    tn = normalize(target_col)
    found = [tc for tc_n, tc in template_norm.items() if tc_n == tn]
    if found:
        out[found[0]] = merged[src_col].values
        out_filled.add(found[0])
        log(f"    CONCAT.{src_col:40s} -> {found[0][:55]}")
    else:
        log(f"    CONCAT.{src_col:40s} -> NO ENCONTRADO en template")

# ---- MAP UDS FIELDS (fill where CONCAT didn't) ----
uds_mapping = {
    "EntidadContratista": "CONTRATISTA 2026",
    "DepartamentoUDS": "Departamento UDS",
    "BarrioUDS": "Barrio UDS",
    "DireccionUDS": "Direccion UDS",
    "EstadoUDS": "Estado UDS",
    "VigenciaServicio": "Vigencia UDS",
    "CuposUDS": "Cupos UDS ofertados 2025",
    "NombrePropiedadInfraestructura": "Propiedad de\nla infraestructura",
    "ZonaUbicacionUDS": "Zona Ubicación UDS",
    "TipoOrganizacionEC": "Tipo de Organización",
    "NumeroContrato": "Número Contrato",
    "CodigoRegionalUDS": "CodigoRegional",
    "DepartamentoEC": "Departamento Entidad Contratista\n2025",
    "MunicipioEC": "Municipio Entidad Contratista\n2025",
    "CodigoMunicipioEC": "Codigo Municipio Entidad Contratista\n2025",
    "CuposServicioUDS": None,
    "CodigoDepartamentoUDS": None,
    "CodigoMunicipioUDS": None,
    "CodigoBarrioUDS": None,
    "ComunaUDS": None,
    "CentroPobladoUDS": None,
}

log("\n  Mapeando campos desde UDS...")
for src_col, target_col in uds_mapping.items():
    if src_col not in merged.columns:
        continue
    if target_col is None:
        continue
    tn = normalize(target_col)
    found = [tc for tc_n, tc in template_norm.items() if tc_n == tn]
    if found:
        tc = found[0]
        if tc not in out_filled:
            out[tc] = merged[src_col].values
            out_filled.add(tc)
            log(f"    UDS.{src_col:35s} -> {tc[:55]}")
        else:
            # Already filled by CONCAT, only fill where CONCAT has null
            concat_val = merged[concat_mapping.get(src_col, None) or ""] if any(k == src_col for k, v in concat_mapping.items() if v and normalize(v) == normalize(tc)) else None
            # Simple approach: fill only where template col is null
            mask = out[tc].isna()
            if mask.any():
                out.loc[mask, tc] = merged.loc[mask, src_col].values
                log(f"    UDS.{src_col:35s} -> {tc[:55]} (rellenando {mask.sum()} vacios)")
    else:
        log(f"    UDS.{src_col:35s} -> NO ENCONTRADO en template")

# ---- Special: Servicio 2025 (from UDS where CONCAT doesn't fill) ----
tn_servicio2025 = normalize("Servicio 2025")
found_servicio2025 = [tc for tc_n, tc in template_norm.items() if tc_n == tn_servicio2025]
if found_servicio2025:
    tc = found_servicio2025[0]
    if tc not in out_filled:
        out[tc] = merged["Servicio"].values
        out_filled.add(tc)
        log(f"    * Servicio -> {tc[:55]}")
    else:
        mask = out[tc].isna()
        if mask.any():
            out.loc[mask, tc] = merged.loc[mask, "Servicio"].values
            log(f"    * Servicio -> {tc[:55]} (rellenando {mask.sum()} vacios)")

# ---- Special: NIT fields ----
tn_nit_entidad = normalize("NIT_EntidadContratista")
found_nit_entidad = [tc for tc_n, tc in template_norm.items() if tc_n == tn_nit_entidad]
tn_nit_contratista = normalize("NIT CONTRATISTA 2026")
found_nit_contratista = [tc for tc_n, tc in template_norm.items() if tc_n == tn_nit_contratista]

if found_nit_contratista:
    tc = found_nit_contratista[0]
    merged["_nit_clean"] = merged["NumeroDocumentoEC"].apply(clean_nit)
    out[tc] = merged["_nit_clean"].values
    out_filled.add(tc)
    log(f"    * NIT (NumeroDocumentoEC) -> {tc[:55]}")

if found_nit_entidad and found_nit_entidad[0] not in out_filled:
    tc = found_nit_entidad[0]
    merged["_nit_clean2"] = merged["NumeroDocumentoEC"].apply(clean_nit)
    out[tc] = merged["_nit_clean2"].values
    out_filled.add(tc)
    log(f"    * NIT (NumeroDocumentoEC) -> {tc[:55]}")

# ---- Special: Tipo_Modalidad -> Modalidad 2026 ----
tn_modalidad = normalize("Modalidad 2026")
found_modalidad = [tc for tc_n, tc in template_norm.items() if tc_n == tn_modalidad]
if found_modalidad:
    tc = found_modalidad[0]
    out[tc] = merged["Tipo_Modalidad"].values
    out_filled.add(tc)
    log(f"    CONCAT.Tipo_Modalidad -> {tc[:55]}")

# ---- Special: "unive" column (index 0) ----
# This seems to be a row counter, fill with sequential number
if "unive" in TEMPLATE_COLS and "unive" not in out_filled:
    out["unive"] = range(1, len(merged) + 1)
    out_filled.add("unive")
    log(f"    * Rellenando 'unive' con numeracion secuencial")

# ---- Special: LITERAL_DE_CONTRATACION -> also try TIPO DE\nCONTRATACIÓN 2026-SUGERIDO ----
tn_literal_sug = normalize("TIPO DE\nCONTRATACIÓN 2026-SUGERIDO ÁREA TECNICA")
found_literal_sug = [tc for tc_n, tc in template_norm.items() if tc_n == tn_literal_sug]
if found_literal_sug:
    tc = found_literal_sug[0]
    if tc not in out_filled:
        out[tc] = merged["LITERAL_DE_CONTRATACION"].values
        out_filled.add(tc)
        log(f"    CONCAT.LITERAL_DE_CONTRATACION -> {tc[:55]}")

# ============================================================
log("PASO 6: Agregar columnas extra de UDS al final")
# ============================================================
# Columns from UDS that were not mapped into template -> add at end
extra_cols = []
uds_used_in_template = set()
for v in uds_mapping.values():
    if v:
        tn = normalize(v)
        found = [tc for tc_n, tc in template_norm.items() if tc_n == tn]
        if found:
            uds_used_in_template.add(v)

for c in uds.columns:
    if c == "_k":
        continue
    if c.startswith("_k"):
        continue
    if c.startswith("_"):
        continue
    # Check if this column was already mapped to template
    already_mapped = False
    for src, tgt in uds_mapping.items():
        if src == c and tgt is not None:
            already_mapped = True
            break
    if not already_mapped:
        for src, tgt in concat_mapping.items():
            if src == c and tgt is not None:
                already_mapped = True
                break
    if not already_mapped:
        extra_cols.append(c)

log(f"  Columnas extra de UDS a agregar: {len(extra_cols)}")
extra_df = merged[extra_cols].copy() if extra_cols else pd.DataFrame(index=merged.index)
# Rename to prefix with "UDS_" to avoid confusion
extra_df.columns = ["UDS_" + c for c in extra_df.columns]

# ============================================================
log("PASO 7: Armar output final")
# ============================================================
# Start with template columns (filled or not)
final = out.copy()

# Add extra columns at the end
for c in extra_df.columns:
    final[c] = extra_df[c].values

log(f"  Columnas finales: {len(final.columns)}")
log(f"  Columnas llenas (con datos): {(final.notna().sum() > 0).sum()}")

# ============================================================
log("PASO 8: Guardar")
# ============================================================
final.to_excel(OUTPUT_FILE, index=False)
log(f"  Archivo guardado: {OUTPUT_FILE}")
log(f"  {len(final)} filas x {len(final.columns)} columnas")

# ============================================================
log("RESUMEN DE LLENADO")
# ============================================================
print()
print("=" * 70)
print("  COLUMNA                      |  LLENAS  |  VACIAS  |  %")
print("=" * 70)
for c in final.columns:
    llenas = final[c].notna().sum()
    vacias = final[c].isna().sum()
    pct = llenas / len(final) * 100
    if pct > 0:
        print("  %-35s | %8d | %8d | %5.1f%%" % (c[:35], llenas, vacias, pct))
print("=" * 70)

print()
print("PROCESO COMPLETADO")
