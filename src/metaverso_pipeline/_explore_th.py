import pandas as pd, warnings
warnings.filterwarnings('ignore')

PATH = r'data\insumos metaverso\UDS_15052026_3112025.xlsx'
xls = pd.ExcelFile(PATH)
print('Sheets:', xls.sheet_names)

df = pd.read_excel(PATH, sheet_name='TH_Transito_15052026')
print('Shape:', df.shape)
print('Columnas:', len(df.columns))
for i, c in enumerate(df.columns):
    print(f'  [{i}] {repr(c)}')
print()
print('Primeras 3 filas:')
print(df.head(3).to_string())
print()
print('Tipos:')
print(df.dtypes)
