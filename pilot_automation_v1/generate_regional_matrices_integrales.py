import os
import re
import time
import win32com.client as win32

# ==============================================================================
# CONFIGURATION
# ==============================================================================
BASE_DIR = r"D:\ICBF\cost-tracking\data\insumos 1 julio"
TEMPLATE_FILE = os.path.join(
    BASE_DIR,
    "Plantilla_Matriz_Adición_Nivelacion_ V2_20032026_paradividir_correcion.xlsx"
)
REGIONAL_DATA_DIR = os.path.join(
    BASE_DIR, "Zonificación_IIsemestre_Integrales"
)
OUTPUT_DIR = r"D:\ICBF\cost-tracking\data\matrices_integrales_por_regional"

PASSWORD = "#APOLO1704*"

SHEETS_TO_KEEP = {
    "MATRIZ_CALCULADA",
    "ZONIFICACIÓN- PEDAGOGICO",
    "DILIGENCIAR",
    "Parametros",
    "ESCALA SALARIAL",
    "RP",
    "BASE_TASAS",
}

COLUMN_MAPPING = {
    1: 0,  # ZONA 2026.           <- Número de Zona
    2: 6,  # Regional UDS          <- Regional UDS
    3: 10, # Centro Zonal UDS      <- Centro Zonal UDS
    4: 8,  # Municipio UDS         <- Municipio UDS
    5: 22, # SERVICIO 2026         <- SERVICIO 2026 II Semestre
    6: 34, # Componente para UDS   <- Componente 2026 II Semestre
    7: 12, # Codigo Unidad Servicio UDS <- Codigo Unidad Servicio UDS
    8: 13, # Unidad Servicio UDS   <- Unidad Servicio UDS
    9: 25, # Cupos                 <- Cupos a Programar 2026 II Semestre
}

PED_DATA_COLS = set(COLUMN_MAPPING.keys())


def clean_name(name):
    name = str(name).strip().upper()
    for a, b in [('Á','A'),('É','E'),('Í','I'),('Ó','O'),('Ú','U'),('Ñ','N')]:
        name = name.replace(a, b)
    name = name.replace(' ', '_').replace('.', '')
    return re.sub(r'_+', '_', name).strip('_')


def find_regional_file(regional_name):
    folder = os.path.join(REGIONAL_DATA_DIR, regional_name)
    if not os.path.isdir(folder):
        return None
    candidates = [
        f for f in os.listdir(folder)
        if f.endswith('.xlsm') and 'Base' in f
    ]
    if not candidates:
        return None
    # Prefer file without "ok" suffix, then longest name
    candidates.sort(key=lambda x: (x.endswith('ok.xlsm'), len(x)), reverse=True)
    return os.path.join(folder, candidates[0])


def find_zonificacion_sheet(wb):
    for s in wb.Sheets:
        name = s.Name
        if "Zonificación" in name and "(2)" not in name:
            return s
    return wb.Sheets(1)


def read_regional_data(excel, regional_file):
    wb = excel.Workbooks.Open(regional_file, UpdateLinks=0, ReadOnly=True)
    ws = find_zonificacion_sheet(wb)
    last_row = ws.UsedRange.Rows.Count
    last_col = ws.UsedRange.Columns.Count
    actual_last = 1
    for r in range(last_row, 1, -1):
        if ws.Cells(r, 1).Value is not None:
            actual_last = r
            break
    rows = []
    for r in range(2, actual_last + 1):
        row = []
        for c in range(1, last_col + 1):
            row.append(ws.Cells(r, c).Value)
        rows.append(row)
    wb.Close()
    return rows



def fix_pedagogico_cost_formulas(ws_ped, num_data_rows):
    """Wrap VLOOKUPs in cost columns (L, M, N) with IFERROR to avoid #VALUE!."""
    start_row = 4
    end_row = start_row + num_data_rows - 1
    if num_data_rows == 0:
        return

    for col_idx in [12, 13, 14]:
        f = ws_ped.Cells(start_row, col_idx).Formula
        if not f or not str(f).startswith("="):
            continue
        old = str(f)
        if "IFERROR" in old:
            continue
        if "[@Cupos]" not in old:
            continue
        # Wrap the VLOOKUP(*[@Cupos]) portion with IFERROR(...,0)
        new_f = old
        for vcol in [95, 94, 93]:
            pattern = f"VLOOKUP([@[SERVICIO 2026]],Parametros!$A$3:$CQ$46,{vcol},0)*[@Cupos]"
            replacement = f"IFERROR(VLOOKUP([@[SERVICIO 2026]],Parametros!$A$3:$CQ$46,{vcol},0)*[@Cupos],0)"
            new_f = new_f.replace(pattern, replacement)
        if new_f == old:
            continue
        # Enter modified formula in first row and fill down
        ws_ped.Cells(start_row, col_idx).Formula = new_f
        fill_range = ws_ped.Range(
            ws_ped.Cells(start_row, col_idx),
            ws_ped.Cells(end_row, col_idx)
        )
        fill_range.FillDown()


def prepare_template(wb):
    if PASSWORD:
        try:
            wb.Unprotect(PASSWORD)
        except Exception:
            pass
        for s in wb.Sheets:
            try:
                s.Unprotect(PASSWORD)
            except Exception:
                pass

    for s in list(wb.Sheets):
        name = s.Name
        if name not in SHEETS_TO_KEEP:
            try:
                s.Visible = -1
                s.Delete()
            except Exception:
                pass

    ws_ped = wb.Sheets("ZONIFICACIÓN- PEDAGOGICO")
    ws_mat = wb.Sheets("MATRIZ_CALCULADA")
    return ws_ped, ws_mat


def fill_template(ws_ped, ws_mat, data_rows):
    num_rows = len(data_rows)
    if num_rows == 0:
        return 0

    ped_max = ws_ped.UsedRange.Rows.Count
    if ped_max < 4:
        ped_max = 23783

    # Clear old data cols (A-I)
    for col in PED_DATA_COLS:
        ws_ped.Range(
            ws_ped.Cells(4, col),
            ws_ped.Cells(ped_max, col)
        ).ClearContents()

    # Write new data to A-I
    end_row = 3 + num_rows
    for target_col, source_idx in COLUMN_MAPPING.items():
        col_values = []
        for row_data in data_rows:
            val = row_data[source_idx] if source_idx < len(row_data) else None
            col_values.append(val if val is not None else "")
        col_2d = [[v] for v in col_values]
        ws_ped.Range(
            ws_ped.Cells(4, target_col),
            ws_ped.Cells(end_row, target_col)
        ).Value = col_2d

    # Fix VLOOKUP formulas to prevent #VALUE! for missing services
    fix_pedagogico_cost_formulas(ws_ped, num_rows)

    # Clear stale UNIQUE spill in MATRIZ_CALCULADA
    mat_max = ws_mat.UsedRange.Rows.Count
    if mat_max > 3:
        ws_mat.Range(
            ws_mat.Cells(4, 1),
            ws_mat.Cells(mat_max, 44)
        ).ClearContents()

    return num_rows


def main():
    total_start = time.time()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    regionals = sorted([
        d for d in os.listdir(REGIONAL_DATA_DIR)
        if os.path.isdir(os.path.join(REGIONAL_DATA_DIR, d))
    ])
    print(f"Found {len(regionals)} regional folders\n", flush=True)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False

    done = 0
    try:
        for i, regional in enumerate(regionals, 1):
            reg_start = time.time()
            print(f"[{i}/{len(regionals)}] {regional}...", end=" ", flush=True)

            regional_file = find_regional_file(regional)
            if regional_file is None:
                print("SKIP (no Base_*.xlsm)")
                continue

            try:
                data_rows = read_regional_data(excel, regional_file)
            except Exception as e:
                print(f"ERROR reading: {e}", flush=True)
                continue

            if not data_rows:
                print("SKIP (0 rows)")
                continue

            folder_name = clean_name(regional)
            out_dir = os.path.join(OUTPUT_DIR, folder_name)
            os.makedirs(out_dir, exist_ok=True)
            out_file = os.path.join(out_dir, f"MATRIZ_2026_{folder_name}_INTEGRALES.xlsx")

            try:
                wb = excel.Workbooks.Open(TEMPLATE_FILE, UpdateLinks=0, ReadOnly=False)
                ws_ped, ws_mat = prepare_template(wb)
                num_written = fill_template(ws_ped, ws_mat, data_rows)
                wb.SaveAs(out_file, FileFormat=51)
                wb.Close(False)

                elapsed = time.time() - reg_start
                print(f"OK ({num_written} rows, {elapsed:.1f}s)", flush=True)
                done += 1

            except Exception as e:
                print(f"ERROR: {e}", flush=True)
                try:
                    wb.Close(False)
                except Exception:
                    pass

    finally:
        excel.Quit()

    total = time.time() - total_start
    print(f"\nDone: {done} matrices in {total:.1f}s")


if __name__ == "__main__":
    main()
