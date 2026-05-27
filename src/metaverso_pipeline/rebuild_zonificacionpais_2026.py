import pandas as pd
import os
import re
import unicodedata
import warnings
warnings.filterwarnings('ignore')

DIR_BASE = r"D:\ICBF\cost-tracking\data\insumos metaverso"
CONCAT  = os.path.join(DIR_BASE, "CONCAT_ZONIFICACION_METAVERSO_2026.xlsx")
UDS     = os.path.join(DIR_BASE, "UDS_15052026_3112025.xlsx")
ORIG    = os.path.join(DIR_BASE, "Metaverso 2026.xlsx")
OUTPUT  = os.path.join(DIR_BASE, "ZONIFICACIONPAIS_RECONSTRUIDA_2026.xlsx")

def normalize(s):
    return unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('ASCII').lower().replace(' ', '')

def clean_nit(x):
    if pd.isna(x): return None
    s = str(x).replace(',', '.').split('.')[0]
    s = ''.join(c for c in s if c.isdigit())
    return s if s else None

def col_letter(idx):
    n = idx
    result = []
    while True:
        result.append(chr(65 + n % 26))
        n = n // 26 - 1
        if n < 0: break
    return ''.join(reversed(result))

print("="*50)
print("PASO 1: CONCAT")
df = pd.read_excel(CONCAT)
df["k"] = df["Codigo_UDS"].astype(str).str.strip()
print(f"  {len(df)} filas")

print("PASO 2: UDS_15052026")
uds = pd.read_excel(UDS, sheet_name="UDS_15052026")
uds["k"] = uds["CodigoUnidadServicioUDS"].astype(str).str.strip()
uds = uds.drop_duplicates(subset="k", keep="first")
df = df.merge(uds, on="k", how="left", suffixes=("_cat","_uds"))
print(f"  {df['CodigoUnidadServicioUDS'].notna().sum()} enlazados")

print("PASO 3: TH_Transito_15052026")
th = pd.read_excel(UDS, sheet_name="TH_Transito_15052026")
th["k"] = th["codigounidadservicio"].astype(str).str.strip()
th = th.drop_duplicates(subset="k", keep="first")
df = df.merge(th[["k","total"]].rename(columns={"total":"TH_Transito_Cupos"}), on="k", how="left")
print(f"  {df['TH_Transito_Cupos'].notna().sum()} enlazados")

print("PASO 4: UDS_31122025 (CuposUDS_31122025, CuposServicioUDS_31122025)")
uds25 = pd.read_excel(UDS, sheet_name="UDS_31122025")
uds25["k"] = uds25["CodigoUnidadServicioUDS"].astype(str).str.strip()
uds25 = uds25.drop_duplicates(subset="k", keep="first")
df = df.merge(uds25[["k","CuposUDS_31122025","CuposServicioUDS_31122025"]], on="k", how="left")
print(f"  CuposUDS_31122025: {df['CuposUDS_31122025'].notna().sum()}")
print(f"  CuposServicioUDS_31122025: {df['CuposServicioUDS_31122025'].notna().sum()}")

print("PASO 5: Estructura ZonificacionPais")
meta = pd.read_excel(ORIG, sheet_name="ZonificacionPais")
MONETARY = {" APALANCAMIENTO NUEVO DOS DIAS "," VALOR TOTAL 2026 NUEVO "," VALOR SOLO 2026 ","APALANCA MAS VIGENCIA NUEVO 2026","Costo total hasta Julio ajuste operador alimentos"}
meta_cols = [c for c in list(meta.columns) if c not in MONETARY]
print(f"  {len(meta_cols)} columnas no monetarias")

print("PASO 6: Campos derivados")
df["NIT_EC"] = df["NumeroDocumentoEC"].apply(clean_nit)
df["NumeroContrato_clean"] = df["NumeroContrato"].astype(str).str.strip()
AM = {"HCB":"FAMILIAR Y COMUNITARIA","FAMI":"FAMILIAR Y COMUNITARIA","HCB_SA":"FAMILIAR Y COMUNITARIA","BV":"INSTITUCIONAL","JC":"INSTITUCIONAL"}
def mod(row):
    if pd.notna(row["Abrev"]) and row["Abrev"] in AM: return AM[row["Abrev"]]
    s = str(row["Servicio"]).upper() if pd.notna(row["Servicio"]) else ""
    if "PROPIA" in s or "INTERCULTURAL" in s: return "PROPIA E INTERCULTURAL"
    if "INSTITUCIONAL" in s or "CDI" in s or "CENTRO DE DESARROLLO" in s: return "INSTITUCIONAL"
    if "MODELO PROPIO" in s or "MAI" in s: return "MODELO PROPIO"
    if "FAMILIAR" in s or "COMUNITARIA" in s or "HCB" in s: return "FAMILIAR Y COMUNITARIA"
    return row["Abrev"] if pd.notna(row["Abrev"]) else None
df["Modalidad_calc"] = df.apply(mod, axis=1).fillna(df["Tipo_Modalidad"])

print("PASO 7: Mapear columnas base")
M = {
    "Regional UDS":"Regional_UDS","Centro Zonal UDS":"Centro_Zonal_UDS","Municipio UDS":"Municipio_UDS",
    "ZONA 2026":"ZONA","Codigo Unidad Servicio UDS":"Codigo_UDS","Unidad Servicio UDS":"UDS",
    "SERVICIO\n2026":"Servicio","Cupos a Programar 2026":"Cupos",
    "Cantidad Madres\nComunitarias en la UDS":"Madres_Unds","COMPONENTE":"Componente_para_la_UDS",
    "TIPO DE\nCONTRATACI\u00d3N 2026":"LITERAL_DE_CONTRATACION","Departamento UDS":"DepartamentoUDS",
    "Barrio UDS":"BarrioUDS","Direccion UDS":"DireccionUDS","Estado UDS":"EstadoUDS",
    "Cupos UDS ofertados 2025":"CuposUDS","CONTRATISTA 2026":"EntidadContratista",
    "NIT CONTRATISTA 2026":"NIT_EC","Propiedad de\nla infraestructura ":"NombrePropiedadInfraestructura",
    "Departamento Entidad Contratista\n2025":"DepartamentoEC","Municipio Entidad Contratista\n2025":"MunicipioEC",
    "Codigo Municipio Entidad Contratista\n2025":"CodigoMunicipioEC",
    "Zona Ubicaci\u00f3n UDS":"ZonaUbicacionUDS","Vigencia UDS":"VigenciaServicio",
    "CodigoRegional":"CodigoRegionalUDS","Modalidad 2026":"Modalidad_calc",
    "NIT_EntidadContratista":"EntidadContratista",
    "Modalidad\n2025":"Modalidad_calc",
    "Servicio 2025":"Servicio",
}
out = pd.DataFrame(index=range(len(df)), columns=meta_cols)
mn = {normalize(c):c for c in meta_cols}
asg = set()
for mc, sc in M.items():
    n = normalize(mc)
    if n in mn:
        out[mn[n]] = df[sc].values; asg.add(mn[n])
for c in meta_cols:
    ca = "".join(ch for ch in c if ch.isascii()).lower().strip()
    if "contrato" in ca and ca.startswith("n") and c not in asg:
        out[c] = df["NumeroContrato_clean"].values; asg.add(c); print(f"  Manual: {c}")
print(f"  Mapeadas: {len(asg)}")

print("PASO 8: Agregar 3 nuevas columnas (CuposUDS_31122025, CuposServicioUDS_31122025, TH_Transito_Cupos)")
out["CuposUDS_31122025"] = df["CuposUDS_31122025"].values
out["CuposServicioUDS_31122025"] = df["CuposServicioUDS_31122025"].values
out["TH_Transito_Cupos"] = df["TH_Transito_Cupos"].values
print(f"  Columnas finales: {len(out.columns)}")

# NIT clean (skip NIT_EntidadContratista which stores name)
for c in out.columns:
    if "NIT_EntidadContratista" in str(c): continue
    ca = "".join(ch for ch in c if ch.isascii()).lower().strip()
    if "nit" in ca and "contratista" in ca:
        out[c] = out[c].apply(lambda x: str(int(x)) if pd.notna(x) and str(x).replace(".0","").isdigit() else (str(x) if pd.notna(x) else None))

print("PASO 9: Formulas en Unnamed: 101")
# Find Unnamed: 101 and the two referenced columns
formula_col_idx = None
riesgo_idx = None
servicio_aj_idx = None
for i, c in enumerate(out.columns):
    if str(c).startswith("Unnamed"):
        formula_col_idx = i
    elif c == "RIESGO TRANSPARENCIA":
        riesgo_idx = i
    elif c == "Servicio ajustado + nuevas":
        servicio_aj_idx = i
print(f"  Formula col index: {formula_col_idx}")
print(f"  RIESGO TRANSPARENCIA index: {riesgo_idx} -> col {col_letter(riesgo_idx)}")
print(f"  Servicio ajustado + nuevas index: {servicio_aj_idx} -> col {col_letter(servicio_aj_idx)}")

print("GUARDANDO...")
out.to_excel(OUTPUT, index=False)
print(f"  {OUTPUT}")
print(f"  Filas: {len(out)}, Columnas: {len(out.columns)}")

print("="*50)
print("REPORTE DE ANOMALIAS")
print("="*50)
th_codes = set(th["k"].unique())
concat_codes = set(df["k"].unique())
th_miss = len(th_codes - concat_codes)
print(f"  TH_Transito NO enlazados: {th_miss}/{len(th_codes)}")
u25_codes = set(uds25["k"].unique())
u25_miss = len(u25_codes - concat_codes)
print(f"  UDS_31122025 NO enlazados: {u25_miss}/{len(u25_codes)}")
for cn in ["CuposUDS_31122025","CuposServicioUDS_31122025","TH_Transito_Cupos"]:
    if cn in out.columns:
        empty = out[cn].isna().sum()
        print(f"  {cn} vacios: {empty}/{len(out)}")
print()
print("PROCESO COMPLETADO")
print()
print("AGREGANDO FORMULAS...")
