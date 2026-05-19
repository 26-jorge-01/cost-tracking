import pandas as pd
import os

file_path = 'd:/ICBF/cost-tracking/data/insumos 28 abril/UDS_28042026_10AM.xlsx'
df = pd.read_excel(file_path, sheet_name='UDS_28042026_10AM', usecols=['Nombre Servicio'])
print(df['Nombre Servicio'].unique())
