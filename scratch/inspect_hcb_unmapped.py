import pandas as pd
import unicodedata
import re

def normalize_str(s):
    if not isinstance(s, str):
        return str(s)
    s = s.upper().strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s_clean = re.sub(r'[^A-Z0-9]', '', s)
    return s_clean

def is_mapped(c):
    # This is a simplified check based on 01_consolidacion_hcb.py logic
    if 'REGIONAL' in c: return True
    if 'CONTRATOSECOP' in c: return True
    if 'CENTROZONAL' in c: return True
    if 'MUNICIPIO' in c: return True
    if 'NOMBREDELSERVICIO' in c: return True
    if 'MODALIDADDESERVICIO' in c: return True
    if c == 'NIT': return True
    if 'NOMBREEAS' in c: return True
    if 'CUPOSPORUNIDAD' in c: return True
    if 'MADRESPORUDS' in c and '2025' in c: return True
    if 'VALORTOTALINICIAL' in c: return True
    if 'VALORINICIALCONTRAPARTIDA' in c: return True
    if 'VALORADICIONESHISTORICAS' in c: return True
    if 'VALORADICIONOTROSCONCEPTOS' in c: return True
    if 'VALORREDUCCIONESHISTORICAS' in c: return True
    if 'VALORINEJECUCIONES' in c: return True
    if 'VALORACTUALDELCONTRATO' in c: return True
    if 'VALORTOTALDELAADICIONCONTRATO' in c: return True
    if 'VALORTOTALDELAADICIONAPORTEICBF' in c: return True
    if 'VALORCONTRAPARTIDAADICION' in c: return True
    if 'VALORACTUAL' in c: return True
    if 'VALORTOTALDELAADICION' in c: return True
    if 'VALORFINAL' in c: return True
    if 'VALORADICIONNIVELACIONCANASTA' in c: return True
    if 'SUMATORIADELOSCUPOSDELCONTRATOAADICIONAR' in c: return True
    return False

file_path = r'D:\ICBF\cost-tracking\data\insumos 28 abril\Elkin\HCB\CALDAS_MATRIZ_ADICIONES_HCB_NIVELACION_CANASTA_30032026V2.2.xlsm'

try:
    xls = pd.ExcelFile(file_path)
    sheet_name = 'MATRIZ'
    df_temp = pd.read_excel(xls, sheet_name=sheet_name, nrows=20, header=None)
    h_idx = None
    for idx, row in df_temp.iterrows():
        if any(isinstance(val, str) and 'REGIONAL' == normalize_str(val) for val in row.values):
            h_idx = idx
            break
            
    if h_idx is not None:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=h_idx)
        print("Unmapped columns:")
        for col in df.columns:
            norm = normalize_str(col)
            if not is_mapped(norm):
                print(f"- {col}")
    else:
        print("Could not find header row.")

except Exception as e:
    print(f"Error: {e}")
