import openpyxl
import os

file_path = 'd:/ICBF/cost-tracking/data/insumos 28 abril/Alba/BOYACA/Boyaca_Matriz_Adición_Nivelacion_V3_01042026.xlsx'

if os.path.exists(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=False)
    ws = wb['MATRIZ_CALCULADA']
    
    print("--- INSPECTING MATRIZ_CALCULADA FORMULAS ---")
    for row in range(1, 15):
        row_vals = []
        for col in range(1, 15):
            cell = ws.cell(row=row, column=col)
            val = cell.value
            row_vals.append(val)
        print(f"Row {row}: {row_vals}")
else:
    print("File not found")
