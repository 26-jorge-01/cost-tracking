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
        print(f"Opening workbook: {xlsx_file} for editing...", flush=True)
        wb = excel.Workbooks.Open(xlsx_file, UpdateLinks=0, ReadOnly=False)
        try:
            for sheet in wb.Sheets:
                shapes_count = sheet.Shapes.Count
                if shapes_count > 0:
                    print(f"\nScanning sheet '{sheet.Name}' (has {shapes_count} shapes):", flush=True)
                    # Loop in reverse order to avoid indexing issues when deleting shapes
                    for i in range(shapes_count, 0, -1):
                        shape = sheet.Shapes.Item(i)
                        shape_name = shape.Name
                        shape_type = shape.Type
                        on_action = ""
                        try:
                            on_action = shape.OnAction
                        except Exception:
                            pass
                        
                        # Decide whether to delete
                        # Type 8: msoFormControl (buttons, forms)
                        # Type 12: msoOLEControlObject (ActiveX controls)
                        # or if there is a macro assigned (on_action is not empty)
                        should_delete = False
                        if shape_type in [8, 12]:
                            should_delete = True
                        elif on_action:
                            should_delete = True
                        
                        if should_delete:
                            print(f"  -> Deleting Shape {i}: Name='{shape_name}', Type={shape_type}, OnAction='{on_action}'", flush=True)
                            shape.Delete()
                        else:
                            print(f"  -> Keeping Shape {i}: Name='{shape_name}', Type={shape_type}", flush=True)
            
            print("Saving changes...", flush=True)
            wb.Save()
            print("File saved successfully.", flush=True)
        finally:
            wb.Close(SaveChanges=True)
    except Exception as e:
        print(f"An error occurred: {e}", flush=True)
    finally:
        print("Quitting Excel Application...", flush=True)
        excel.Quit()

if __name__ == "__main__":
    main()
