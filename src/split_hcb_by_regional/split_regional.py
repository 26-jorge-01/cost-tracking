import win32com.client as win32
import os
import re
import shutil
import time


SHEETS_TO_KEEP = {"MATRIZ", "rp", "COSTEO", "servicios homologados", "LISTAS", "Hoja5", "INTEGRALIDAD"}
PASSWORD = "GMYC2026**"


def sanitize_folder_name(name: str) -> str:
    replacements = {
        ' ': '_', '.': '', 'ñ': 'n', 'Ñ': 'N',
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
    }
    for k, v in replacements.items():
        name = name.replace(k, v)
    return re.sub(r'_+', '_', name).strip('_').upper()


def get_regionals(source_path: str, sheet_name: str = "MATRIZ", col: int = 3) -> list[str]:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    try:
        wb = excel.Workbooks.Open(source_path)
        ws = wb.Sheets(sheet_name)
        last_row = ws.UsedRange.Rows.Count
        regionals = set()
        for r in range(3, last_row + 1):
            val = ws.Cells(r, col).Value
            if val:
                regionals.add(str(val).strip())
        wb.Close()
        return sorted(regionals)
    finally:
        excel.Quit()


def _clean_sheets(wb, sheet_names: list[str]):
    for s_name in sheet_names:
        if s_name not in SHEETS_TO_KEEP:
            try:
                wb.Sheets(s_name).Delete()
            except Exception:
                pass


def _unlist_tables(ws):
    while ws.ListObjects.Count > 0:
        ws.ListObjects(1).Unlist()


def _fix_formula_references(ws, last_col: int):
    tlr = ws.UsedRange.Rows.Count
    used_range = ws.Range(ws.Cells(1, 1), ws.Cells(tlr, last_col))
    used_range.Replace(What="MATRIZ!", Replacement="", LookAt=2)


def _filter_and_keep_regional(ws, last_col: int, regional: str):
    tlr = ws.UsedRange.Rows.Count
    data_range = ws.Range(ws.Cells(2, 1), ws.Cells(tlr, last_col))
    data_range.AutoFilter(Field=3, Criteria1="<>" + regional)
    visible = ws.Range(ws.Cells(3, 1), ws.Cells(tlr, last_col)).SpecialCells(12)
    for i in range(visible.Areas.Count, 0, -1):
        visible.Areas(i).EntireRow.Delete()
    if ws.AutoFilterMode:
        ws.AutoFilterMode = False


def _protect_sheets(wb, editable_sheet: str, password: str):
    for s in wb.Sheets:
        if s.Name != editable_sheet:
            s.Protect(Password=password)


def _create_template(source_path: str, sheets_to_keep: set[str]) -> str:
    template_dir = os.path.join(os.path.dirname(source_path), "_template")
    os.makedirs(template_dir, exist_ok=True)
    template_path = os.path.join(template_dir, "_TEMPLATE.xlsx")
    shutil.copy2(source_path, template_path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    try:
        wb = excel.Workbooks.Open(template_path)
        all_sheets = [s.Name for s in wb.Sheets]
        for s_name in all_sheets:
            if s_name not in sheets_to_keep:
                try:
                    wb.Sheets(s_name).Delete()
                except Exception:
                    pass
        ws = wb.Sheets("MATRIZ")
        while ws.ListObjects.Count > 0:
            ws.ListObjects(1).Unlist()
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()
    return template_path


def split_by_regional(
    source_path: str,
    output_root: str,
    sheet_name: str = "MATRIZ",
    regionals: list[str] | None = None,
    sheets_to_keep: set[str] | None = None,
    password: str | None = PASSWORD,
) -> list[str]:
    if sheets_to_keep is None:
        sheets_to_keep = SHEETS_TO_KEEP
    os.makedirs(output_root, exist_ok=True)

    if regionals is None:
        regionals = get_regionals(source_path, sheet_name)

    print("Preparando template optimizado...", end=" ", flush=True)
    t_tmpl = time.time()
    template_path = _create_template(source_path, sheets_to_keep)
    print(f"OK ({time.time()-t_tmpl:.1f}s)", flush=True)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        created = []
        for i, regional in enumerate(regionals, 1):
            folder_name = sanitize_folder_name(regional)
            folder_path = os.path.join(output_root, folder_name)
            os.makedirs(folder_path, exist_ok=True)
            out_path = os.path.join(folder_path, f"MATRIZ_{folder_name}.xlsx")

            print(f"[{i}/{len(regionals)}] {regional}...", end=" ", flush=True)
            t0 = time.time()

            shutil.copy2(template_path, out_path)
            wb_copy = excel.Workbooks.Open(out_path)
            ws_copy = wb_copy.Sheets(sheet_name)

            lr = ws_copy.UsedRange.Rows.Count
            lc = ws_copy.UsedRange.Columns.Count

            ws_temp = wb_copy.Sheets.Add()
            ws_temp.Name = "TEMP"
            ws_copy.Range(ws_copy.Cells(1, 1), ws_copy.Cells(lr, lc)).Copy(Destination=ws_temp.Cells(1, 1))

            _fix_formula_references(ws_temp, lc)
            _filter_and_keep_regional(ws_temp, lc, regional)

            ws_copy.Delete()
            ws_temp.Name = sheet_name

            if password:
                _protect_sheets(wb_copy, sheet_name, password)

            wb_copy.Save()
            wb_copy.Close()
            created.append(out_path)
            print(f"OK ({time.time()-t0:.1f}s)", flush=True)

        return created
    finally:
        excel.Quit()
        try:
            os.remove(template_path)
            template_dir = os.path.dirname(template_path)
            if not os.listdir(template_dir):
                os.rmdir(template_dir)
        except Exception:
            pass


def main():
    source = r"d:\ICBF\cost-tracking\data\replicacion hcb 5 junio\MATRIZ_2026_ADICION_COMUNITARIOS_MODFV6.xlsx"
    output_root = r"d:\ICBF\cost-tracking\data\replicacion hcb 5 junio\REGIONALES"
    t_start = time.time()
    regionals = get_regionals(source)
    print(f"Regionales encontradas: {len(regionals)}")
    print(f"Protegiendo hojas (excepto MATRIZ) con contraseña: {PASSWORD}")
    created = split_by_regional(source, output_root, regionals=regionals, password=PASSWORD)
    print(f"\nProceso completado en {time.time()-t_start:.1f}s — {len(created)} archivos generados")


if __name__ == "__main__":
    main()
