import pandas as pd
import os

file_path = r'D:\ICBF\cost-tracking\data\insumos 28 abril\consolidacion_matriz_hcb_28042026_COPIA_SEGURA.xlsx'

if not os.path.exists(file_path):
    print(f"Error: No se encuentra el archivo {file_path}. Ejecuta los pasos 01 y 02 de HCB.")
    exit()

print(f"Iniciando auditoría HCB sobre: {os.path.basename(file_path)}")

try:
    df = pd.read_excel(file_path, sheet_name='MATRIZ_CALCULADA')
    
    print(f"\n--- ESTADÍSTICAS HCB ---")
    print(f"Total registros consolidados: {len(df)}")
    print(f"Regionales presentes: {len(df['REGIONAL'].unique())}")
    
    # 1. Detección de Totales
    mask_total = df.apply(lambda row: row.astype(str).str.contains('TOTAL', case=False).any(), axis=1)
    total_rows = df[mask_total]
    if not total_rows.empty:
        print(f"\n[ALERTA] Se encontraron {len(total_rows)} filas que parecen ser TOTALES.")
        print(total_rows[['REGIONAL', 'Archivo_Origen']].head())
    
    # 2. Verificación de duplicados por contrato
    # En HCB a veces un contrato tiene muchas filas por diferentes conceptos, 
    # pero revisamos si hay algo sospechoso.
    contract_col = 'REFERENCIA / No. CONTRATO SECOP'
    if contract_col in df.columns:
        counts = df[contract_col].value_counts()
        print(f"\nContratos analizados: {len(counts)}")
        
except Exception as e:
    print(f"Error: {e}")
