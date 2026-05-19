import pandas as pd
import os

# Configuración de ruta
file_path = r'D:\ICBF\cost-tracking\data\insumos 28 abril\consolidacion_matriz_integrales_28042026_COPIA_SEGURA.xlsx'

if not os.path.exists(file_path):
    print(f"Error: No se encuentra el archivo {file_path}. Ejecuta los pasos 01 y 02 primero.")
    exit()

print(f"Iniciando auditoría sobre: {os.path.basename(file_path)}")

try:
    # 1. Cargar datos
    xls = pd.ExcelFile(file_path)
    df_matriz = pd.read_excel(xls, sheet_name='MATRIZ_CALCULADA')
    df_zoni = pd.read_excel(xls, sheet_name='ZONIFICACIÓN- PEDAGOGICO')
    
    print(f"\n--- ESTADÍSTICAS GENERALES ---")
    print(f"Filas en MATRIZ: {len(df_matriz)}")
    print(f"Filas en ZONIFICACIÓN: {len(df_zoni)}")
    
    # 2. Detección de filas de "TOTAL" (Riesgo de doble conteo)
    mask_total = df_matriz.apply(lambda row: row.astype(str).str.contains('TOTAL', case=False).any(), axis=1)
    total_rows = df_matriz[mask_total]
    if not total_rows.empty:
        print(f"\n[ALERTA] Se encontraron {len(total_rows)} filas que parecen ser SUBTOTALES o TOTALES.")
        print("Estas filas podrían estar duplicando los valores si se suman directamente.")
        print(total_rows[['REGIONAL', 'MUNICIPIO', 'VALOR ACTUAL']].head())
    else:
        print("\n[OK] No se detectaron filas de 'TOTAL' evidentes en la Matriz.")

    # 3. Análisis de Contratos Duplicados
    contract_col = 'REFERENCIA / No. CONTRATO SECOP'
    if contract_col in df_matriz.columns:
        counts = df_matriz[contract_col].value_counts()
        duplicados = counts[counts > 1]
        print(f"\n--- ANÁLISIS DE CONTRATOS ---")
        print(f"Contratos con múltiples filas: {len(duplicados)}")
        if len(duplicados) > 0:
            print("Ejemplo de contrato repetido:")
            sample_id = duplicados.index[0]
            print(df_matriz[df_matriz[contract_col] == sample_id][['REGIONAL', 'MUNICIPIO', 'CUPOS', 'VALOR ACTUAL']])

    # 4. Verificación de Cupos
    if 'CUPOS' in df_matriz.columns and 'Cupos_Zona' in df_matriz.columns:
        print(f"\n--- VERIFICACIÓN DE CUPOS ---")
        diff = df_matriz[df_matriz['CUPOS'] != df_matriz['Cupos_Zona']]
        if not diff.empty:
            print(f"[!] Hay {len(diff)} filas donde CUPOS no coincide con Cupos_Zona.")
        else:
            print("[OK] Los cupos coinciden perfectamente.")

except Exception as e:
    print(f"Error durante la auditoría: {e}")
