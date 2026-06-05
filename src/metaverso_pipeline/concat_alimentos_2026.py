import pandas as pd
import os

DIR_BASE = r"D:\ICBF\cost-tracking\data\insumos metaverso"
ALIMENTOS_FILE = r"D:\ICBF\cost-tracking\data\insumos matriz 24 abril\CONTRATOS ALIMENTOS ORGANIZACIONES CAMPESINAS.xlsx"
OUTPUT_FILE = os.path.join(DIR_BASE, "CONCAT_ALIMENTOS_2026.xlsx")

COLUMNS_ALIMENTOS_ZONA = {
    0: "ZONA",
    1: "zona_contratacion",
    2: "Zona_Direccion",
    3: "Verifcada_Dire",
    4: "Regional_UDS",
    5: "Municipio_UDS",
    6: "Centro_Zonal_UDS",
    7: "Codigo_UDS",
    8: "UDS",
    9: "Servicio",
    10: "RESULTADO",
    11: "SERVICIO_SIM",
    12: "Componente_para_la_UDS",
    13: "Cupos",
    14: "CONCEPTO_DEFINITIVO",
    15: "CONTRATISTA",
    16: "NIT_CONTRATISTA",
    17: "OBSERVACION",
}

COLUMNS_ALIMENTOS_GUAJIRA = {
    0: "ZONA",
    1: "Regional_UDS",
    2: "Municipio_UDS",
    3: "Centro_Zonal_UDS",
    4: "Codigo_UDS",
    5: "UDS",
    6: "Servicio",
    7: "SERVICIO_SIM",
    8: "Componente_para_la_UDS",
    9: "Cupos",
    11: "CONCEPTO_DEFINITIVO",
    12: "CONTRATISTA",
    13: "NIT_CONTRATISTA",
}

KEY_COLS = [
    "ZONA", "Regional_UDS", "Municipio_UDS", "Centro_Zonal_UDS",
    "Codigo_UDS", "UDS", "Servicio", "Componente_para_la_UDS",
    "Cupos", "CONCEPTO_DEFINITIVO", "CONTRATISTA", "NIT_CONTRATISTA",
]

def main():
    print(">>> Leyendo ALIMENTOS ZONIFICACION...")
    df_zona = pd.read_excel(ALIMENTOS_FILE, sheet_name="ZONIFICACION", header=None, skiprows=4)
    cols_zona = list(COLUMNS_ALIMENTOS_ZONA.keys())
    df_zona = df_zona.iloc[:, cols_zona].copy()
    df_zona.rename(columns=COLUMNS_ALIMENTOS_ZONA, inplace=True)
    df_zona["Origen"] = "ZONIFICACION"
    print(f"  {len(df_zona)} filas de ZONIFICACION")

    print(">>> Leyendo ALIMENTOS ZONIFICACION LA GUAJIRA...")
    df_guajira = pd.read_excel(ALIMENTOS_FILE, sheet_name="ZONIFICACION LA GUAJIRA", header=None, skiprows=2)
    cols_guajira = list(COLUMNS_ALIMENTOS_GUAJIRA.keys())
    df_guajira = df_guajira.iloc[:, cols_guajira].copy()
    df_guajira.rename(columns=COLUMNS_ALIMENTOS_GUAJIRA, inplace=True)
    df_guajira["Origen"] = "ZONIFICACION LA GUAJIRA"
    print(f"  {len(df_guajira)} filas de ZONIFICACION LA GUAJIRA")

    print(">>> Concatenando y limpiando...")
    df_concat = pd.concat([df_zona, df_guajira], axis=0, ignore_index=True, sort=False)

    available = [c for c in KEY_COLS if c in df_concat.columns]
    extra = [c for c in df_concat.columns if c not in KEY_COLS and c != "Origen"]
    cols_order = available + extra + ["Origen"]
    df_concat = df_concat[[c for c in cols_order if c in df_concat.columns]]

    df_concat["Codigo_UDS"] = df_concat["Codigo_UDS"].astype(str).str.strip().str.replace(".0", "", regex=False)

    print(f">>> Filas totales: {len(df_concat)}")
    print(f">>> Columnas: {list(df_concat.columns)}")

    df_concat.to_excel(OUTPUT_FILE, index=False)
    print(f">>> Archivo guardado: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
