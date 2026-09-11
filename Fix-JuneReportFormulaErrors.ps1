
$ErrorActionPreference = "Stop"
$path = "C:\Users\User\Desktop\Global Purchase Report Claude\REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_2026-06.xlsx"

function Release-Com {
    param($obj)
    if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false

try {
    $wb = $excel.Workbooks.Open($path, [Type]::Missing, $false)
    $excel.Calculation = -4135  # xlCalculationManual - avoid recalculating the whole sheet (incl. live AA/AB MAXIFS/MINIFS) on every step
    $ws = $wb.Worksheets.Item("MHK+STC+MZQ+MIL+RACO Data")
    $lastRow = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
    Write-Host "Last row: $lastRow"

    Write-Host "`n[1] Fixing EURD -> EUR currency typo (N32130, N32321)..."
    Write-Host "  Before: N32130=$($ws.Range('N32130').Value2)  N32321=$($ws.Range('N32321').Value2)"
    $ws.Range("N32130").Value2 = "EUR"
    $ws.Range("N32321").Value2 = "EUR"
    Write-Host "  After:  N32130=$($ws.Range('N32130').Value2)  N32321=$($ws.Range('N32321').Value2)"

    # Recalculate Q/R/S (exchange-rate-dependent) for the two fixed rows, and also AA/AB live formulas
    # for the whole table now that the bad currency is gone (range-scoped, not a full-sheet Calculate()).
    Write-Host "`nRecalculating Q/R/S for the fixed rows, and AA/AB across the table..."
    $ws.Range("Q32130:S32130").Calculate()
    $ws.Range("Q32321:S32321").Calculate()
    $ws.Range("AA2:AB$lastRow").Calculate()
    $wb.Save()
    Write-Host "  Saved (checkpoint 1)."

    foreach ($col in @("AC","AD","AE")) {
        Write-Host "`n[2] Refreshing $col`3:$col$lastRow from row-2 master formula..."
        $master = $ws.Range($col + "2")
        $target = $ws.Range($col + "3:" + $col + $lastRow)
        $master.Copy()
        $target.PasteSpecial(-4123) | Out-Null   # xlPasteFormulas
        Write-Host "  Formulas pasted, calculating range only..."
        $full = $ws.Range($col + "2:" + $col + $lastRow)
        $full.Calculate()
        Write-Host "  Freezing $col`2:$col$lastRow to static values..."
        $full.Copy()
        $full.PasteSpecial(-4163) | Out-Null   # xlPasteValues
        Release-Com $master
        Release-Com $target
        Release-Com $full
        $wb.Save()
        Write-Host "  Saved (checkpoint $col)."
    }

    foreach ($col in @("AG","AH")) {
        Write-Host "`n[3] Refreshing $col`3:$col$lastRow from row-2 master formula..."
        $master = $ws.Range($col + "2")
        $target = $ws.Range($col + "3:" + $col + $lastRow)
        $master.Copy()
        $target.PasteSpecial(-4123) | Out-Null
        Write-Host "  Formulas pasted, calculating range only..."
        $full = $ws.Range($col + "2:" + $col + $lastRow)
        $full.Calculate()
        Write-Host "  Freezing $col`2:$col$lastRow to static values..."
        $full.Copy()
        $full.PasteSpecial(-4163) | Out-Null
        Release-Com $master
        Release-Com $target
        Release-Com $full
        $wb.Save()
        Write-Host "  Saved (checkpoint $col)."
    }

    Write-Host "`n[4] Re-scanning for formula errors..."
    $dataRange = $ws.Range($ws.Cells(1,1), $ws.Cells($lastRow, 36))
    try {
        $errorCells = $dataRange.SpecialCells(-4123, 16)
        Write-Host "STILL HAS ERRORS: $($errorCells.Count) cells"
        $colCounts = @{}
        foreach ($area in $errorCells.Areas) {
            foreach ($cell in $area) {
                $cl = $cell.Address($false,$false) -replace '\d+',''
                if (-not $colCounts.ContainsKey($cl)) { $colCounts[$cl] = 0 }
                $colCounts[$cl]++
            }
        }
        $colCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host "  $($_.Name): $($_.Value)" }
    } catch {
        Write-Host "NO ERRORS FOUND - all clear."
    }

    $excel.Calculation = -4105  # xlCalculationAutomatic - restore default before final save/close
    Write-Host "`n[5] Final save..."
    $wb.Save()
    Write-Host "Saved: $path"

    $wb.Close($false)
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    $excel.Quit()
    Release-Com $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
