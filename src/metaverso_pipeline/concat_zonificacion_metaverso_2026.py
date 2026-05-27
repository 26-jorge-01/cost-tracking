import pandas as pd
import os

DIR_BASE = r"D:\ICBF\cost-tracking\data\insumos metaverso"
INPUT_FILE = os.path.join(DIR_BASE, "zonificación_abastecimiento_servicios_primera_infancia25052026.xlsx")
OUTPUT_FILE = os.path.join(DIR_BASE, "CONCAT_ZONIFICACION_METAVERSO_2026.xlsx")

COLUMNS_COMUNITARIOS = {
    0: "Regional_UDS",
    1: "Centro_Zonal_UDS",
    2: "Municipio_UDS",
    3: "ZONA",
    4: "Codigo_UDS",
    5: "UDS",
    6: "LITERAL_DE_CONTRATACION",
    7: "LITERAL_DE_CONTRATACION_ALTERNATIVO",
    8: "Servicio",
    9: "Atenciones",
    10: "Atenciones_30042026",
    11: "Cupos",
    12: "Madres_Unds",
    13: "Abrev",
}

COLUMNS_INTEGRALES = {
    0: "ZONA",
    1: "Regional_UDS",
    2: "Centro_Zonal_UDS",
    3: "Municipio_UDS",
    4: "Codigo_UDS",
    5: "UDS",
    6: "LITERAL_DE_CONTRATACION",
    7: "LITERAL_DE_CONTRATACION_ALTERNATIVO",
    8: "Servicio",
    9: "Componente_para_la_UDS",
    12: "Cupos",
}

def main():
    print(">>> Leyendo hoja Comunitarios (columnas A–N)...")
    df_com = pd.read_excel(INPUT_FILE, sheet_name="Comunitarios", header=None, skiprows=1)
    df_com = df_com.iloc[:, 0:14].copy()
    df_com.rename(columns=COLUMNS_COMUNITARIOS, inplace=True)
    df_com["Tipo_Modalidad"] = "COMUNITARIO"

    print(">>> Leyendo hoja Integrales-Convenios (columnas A–J + M)...")
    df_int = pd.read_excel(INPUT_FILE, sheet_name="Integrales-Convenios", header=None, skiprows=1)
    cols_int = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 12]
    df_int = df_int.iloc[:, cols_int].copy()
    df_int.rename(columns=COLUMNS_INTEGRALES, inplace=True)
    df_int["Tipo_Modalidad"] = "INTEGRAL"

    common_cols = [
        "Regional_UDS", "Centro_Zonal_UDS", "Municipio_UDS", "ZONA",
        "Codigo_UDS", "UDS",
        "LITERAL_DE_CONTRATACION", "LITERAL_DE_CONTRATACION_ALTERNATIVO",
        "Servicio", "Cupos",
    ]

    print(">>> Concatenando ambas hojas...")
    df_concat = pd.concat([df_com, df_int], axis=0, ignore_index=True, sort=False)

    cols_order = common_cols + [
        c for c in df_concat.columns if c not in common_cols and c != "Tipo_Modalidad"
    ] + ["Tipo_Modalidad"]
    df_concat = df_concat[[c for c in cols_order if c in df_concat.columns]]

    print(f">>> Filas totales: {len(df_concat)}")
    print(f">>> Columnas: {list(df_concat.columns)}")

    df_concat.to_excel(OUTPUT_FILE, index=False)
    print(f">>> Archivo guardado: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
