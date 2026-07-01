import win32com.client as win32

excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AskToUpdateLinks = False
excel.EnableEvents = False

tmpl = r"D:\ICBF\cost-tracking\data\insumos 1 julio\Plantilla_Matriz_Adición_Nivelacion_ V2_20032026_paradividir_correcion.xlsx"
wb = excel.Workbooks.Open(tmpl, UpdateLinks=0, ReadOnly=True)

# Check ALL formulas in MATRIZ_CALCULADA Row 3
ws = wb.Sheets("MATRIZ_CALCULADA")
print("=== MATRIZ_CALCULADA Row 3 - ALL formulas ===")
for c in range(1, 45):
    f = ws.Cells(3, c).Formula
    v = ws.Cells(3, c).Value
    if f:
        f_clean = str(f)[:120]
        print(f"  Col {c}: {f_clean}")
    elif v:
        print(f"  Col {c}: VALUE = {str(v)[:60]}")

# Check ZONIFICACIÓN- PEDAGOGICO Row 4 ALL formulas
ws2 = wb.Sheets("ZONIFICACIÓN- PEDAGOGICO")
print("\n=== ZONIFICACIÓN- PEDAGOGICO Row 4 - ALL formulas ===")
for c in range(1, 18):
    f = ws2.Cells(4, c).Formula
    v = ws2.Cells(4, c).Value
    if f:
        f_clean = str(f)[:120]
        print(f"  Col {c}: {f_clean}")
    elif v:
        print(f"  Col {c}: VALUE = {str(v)[:60]}")

wb.Close()
excel.Quit()
