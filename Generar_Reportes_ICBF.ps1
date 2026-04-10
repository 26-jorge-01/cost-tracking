# =============================================================================
# SCRIPT DE AUTOMATIZACION DE REPORTES ICBF (VERSION DEFINITIVA 9.1 - ASCII)
# =============================================================================

$InputPath = "C:\Users\User\Documents\TRABAJO\ICBF\DESARROLLOS\cost-tracking\data\Matriz_Adicion_Nivelacion_Canasta_ HCB_01042026_V2.xlsm"
$OutputDir = "C:\Users\User\Documents\TRABAJO\ICBF\DESARROLLOS\cost-tracking\data\salidas_definitivas"
$SheetName = "MATRIZ"
$Password = "GMYC2023***"

if (!(Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir }

try {
    Write-Host "--- INICIANDO PROCESO (VERSION 9.1 - SIN ACENTOS) ---" -ForegroundColor White
    
    # 1. Fase de Deteccion (Lectura unica de datos)
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Excel.DisplayAlerts = $False
    $Wb = $Excel.Workbooks.Open($InputPath)
    $Ws = $Wb.Sheets.Item($SheetName)
    
    $ColIdx = 0
    $HeaderRow = 0
    for ($r = 1; $r -le 5; $r++) {
        for ($c = 1; $c -le 10; $c++) {
            if ($Ws.Cells.Item($r, $c).Value2 -eq "REGIONAL") {
                $ColIdx = $c
                $HeaderRow = $r
                break
            }
        }
        if ($ColIdx -ne 0) { break }
    }
    
    if ($ColIdx -eq 0) { throw "No se encontro la columna REGIONAL" }
    
    $LastRow = $Ws.UsedRange.Rows.Count
    $Values = $Ws.Range($Ws.Cells.Item($HeaderRow + 1, $ColIdx), $Ws.Cells.Item($LastRow, $ColIdx)).Value2
    
    $Regionales = @()
    for ($i = 1; $i -le $Values.GetLength(0); $i++) {
        $v = $Values.GetValue($i, 1)
        if ($v -and ($Regionales -notcontains $v)) { $Regionales += $v }
    }
    
    $Wb.Close($False)
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    
    Write-Host "Detectadas $($Regionales.Count) regionales. Procesando..." -ForegroundColor Green

    # 2. Fase de Generacion por Region
    foreach ($Target in $Regionales) {
        Write-Host "-> Generando: $Target" -ForegroundColor Cyan
        $Start = Get-Date
        
        $Instance = New-Object -ComObject Excel.Application
        $Instance.Visible = $False
        $Instance.DisplayAlerts = $False
        
        try {
            $Wb = $Instance.Workbooks.Open($InputPath)
            $Ws = $Wb.Sheets.Item($SheetName)
            try { $Ws.Unprotect($Password) } catch {}

            $Blocks = @()
            $i = 1
            $CountRows = $Values.GetLength(0)
            
            while ($i -le $CountRows) {
                if ($Values.GetValue($i, 1) -ne $Target) {
                    $s = $i + $HeaderRow
                    while ($i -le $CountRows -and $Values.GetValue($i, 1) -ne $Target) { $i++ }
                    $e = $i + $HeaderRow - 1
                    $Blocks += , @($s, $e)
                }
                else {
                    $i++
                }
            }

            if ($Blocks.Count -gt 0) {
                [array]::Reverse($Blocks)
                foreach ($B in $Blocks) {
                    $startR = $B[0]
                    $endR = $B[1]
                    $Ws.Rows("$($startR):$($endR)").Delete() | Out-Null
                }
            }

            if ($Ws.AutoFilterMode) { $Ws.AutoFilterMode = $False }
            $MaxNow = $Ws.UsedRange.Rows.Count
            $Ws.Range($Ws.Cells.Item($HeaderRow, 1), $Ws.Cells.Item($MaxNow, $ColIdx)).AutoFilter()
            
            $SafeName = $Target -replace '[^a-zA-Z0-9]', '_'
            $FileName = "Reporte_$SafeName.xlsm"
            $OutputPath = Join-Path $OutputDir $FileName
            
            $Wb.SaveAs($OutputPath, 52)
            $Wb.Close($False)
            
            $Duration = (Get-Date) - $Start
            Write-Host "   OK: $FileName (Tiempo: $( [Math]::Round($Duration.TotalSeconds, 1) ) seg)" -ForegroundColor Green
            
        }
        finally {
            $Instance.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Instance) | Out-Null
        }
        Start-Sleep -Milliseconds 500
    }

}
catch {
    Write-Error "Error: $($_.Exception.Message)"
}

Write-Host "PROCESO FINALIZADO." -ForegroundColor White
