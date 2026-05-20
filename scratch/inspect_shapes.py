import os
import win32com.client as win32

def main():
    base_dir = r"D:\ICBF\cost-tracking\pilot_automation_v1\master_data"
    xlsx_file = os.path.join(base_dir, "MATRIZ_ADICIONES_HCB_NIVELACION_REGIONAL_TOLIMA_26052025_09CONTV1.xlsx")

    if not os.path.exists(xlsx_file):
        print(f"Error: File not found at: {xlsx_file}")
        return

    print("Initializing Excel COM Application...", flush=True)
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False

    try:
        print(f"Opening workbook: {xlsx_file}", flush=True)
        wb = excel.Workbooks.Open(xlsx_file, UpdateLinks=0, ReadOnly=True)
        try:
            for sheet in wb.Sheets:
                shapes_count = sheet.Shapes.Count
                if shapes_count > 0:
                    print(f"\nSheet '{sheet.Name}' has {shapes_count} shapes:", flush=True)
                    for i in range(1, shapes_count + 1):
                        shape = sheet.Shapes.Item(i)
                        # Try to get Name, Type, and OnAction (macro link)
                        shape_name = shape.Name
                        shape_type = shape.Type
                        on_action = ""
                        try:
                            on_action = shape.OnAction
                        except Exception:
                            pass
                        print(f"  Shape {i}: Name='{shape_name}', Type={shape_type}, OnAction='{on_action}'", flush=True)
        finally:
            wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"An error occurred: {e}", flush=True)
    finally:
        print("Quitting Excel Application...", flush=True)
        excel.Quit()

if __name__ == "__main__":
    main()
