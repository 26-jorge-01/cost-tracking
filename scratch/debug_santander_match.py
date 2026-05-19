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

ALIASES_REGIONALES = {
    'NSANTANDER': 'NORTE DE SANTANDER',
    'NORTEDESANTANDER': 'NORTE DE SANTANDER',
    'N.SANTANDER': 'NORTE DE SANTANDER',
    'SANTANDER': 'SANTANDER',
}

f = "MATRIZ_ADICIONES_HCB_NIVELACION_CANASTA_14042026_V2.xlsm"
root = "D:\\ICBF\\cost-tracking\\data\\Matrices_validadas_definitivas\\Lady\\HCB"

f_norm = normalize_str(f)
root_norm = normalize_str(root)
target_regional = None

for alias, official in ALIASES_REGIONALES.items():
    if alias in f_norm or alias in root_norm:
        target_regional = official
        break

print(f"File: {f}")
print(f"Folder: {root}")
print(f"Target Regional identified: {target_regional}")
