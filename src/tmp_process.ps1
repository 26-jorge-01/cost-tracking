
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open("/app/data/MATRIZ_ADICIONES_HCB_NIVELACION_CANASTA_27032026V2 2.xlsm")
        $ws = $wb.Sheets.Item("MATRIZ")
        $ws.Unprotect("GMYC2023***")

        $lastRow = $ws.UsedRange.Rows.Count
        $range = $ws.Range("A2:BY$lastRow")

        # Encontrar columna REGIONAL dinámicamente
        $colIndex = 1
        $range.AutoFilter($colIndex, "<>ANTIOQUIA")

        try {
            $visibleRange = $ws.Range("A3:A$lastRow").SpecialCells(12)
            if ($visibleRange) { $visibleRange.EntireRow.Delete() }
        } catch { }

        if ($ws.FilterMode) { $ws.ShowAllData() }

        # Refrescar tablas dinámicas
        foreach ($pws in $wb.Worksheets) {
            foreach ($pivot in $pws.PivotTables()) { $pivot.PivotCache().Refresh() }
        }

        $wb.SaveAs("/app/data/salidas/Reporte_ANTIOQUIA.xlsm", 52) 
        $wb.Close($false)
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    