import os
import win32com.client as win32

def main():
    base_dir = r"D:\ICBF\cost-tracking\pilot_automation_v1\master_data"
    src_file = os.path.join(base_dir, "MATRIZ_ADICIONES_HCB_NIVELACION_REGIONAL_TOLIMA_26052025_09CONTV1.xlsm")
    dst_file = os.path.join(base_dir, "MATRIZ_ADICIONES_HCB_NIVELACION_REGIONAL_TOLIMA_26052025_09CONTV1.xlsx")

    if not os.path.exists(src_file):
        print(f"Error: Source file not found at: {src_file}")
        return

    print("Initializing Excel COM Application...", flush=True)
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False

    try:
        print(f"Opening workbook: {src_file}", flush=True)
        # Open workbook (ReadOnly=True, UpdateLinks=0 to avoid prompts)
        wb = excel.Workbooks.Open(src_file, UpdateLinks=0, ReadOnly=True)
        try:
            print(f"Saving workbook as XLSX (Format 51) to: {dst_file}", flush=True)
            # FileFormat=51 represents xlOpenXMLWorkbook (.xlsx) which does not support macros.
            # Excel will automatically strip the VBA project.
            wb.SaveAs(dst_file, FileFormat=51)
            print("Successfully saved file as XLSX.", flush=True)
        finally:
            wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"An error occurred: {e}", flush=True)
    finally:
        print("Quitting Excel Application...", flush=True)
        excel.Quit()

if __name__ == "__main__":
    main()
