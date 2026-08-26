<#
.SYNOPSIS
    Every month, run this once after TVS.xlsx (Finance / Gemini-drafted purchase table) is placed in the TVS folder.

.DESCRIPTION
    1. Refreshes column S (Translated Part Description) formula down to the last data row, so brand-new
       Italian descriptions that have no match in 'DESC ITY - ENG' fall back to Excel's native TRANSLATE().
    2. Counts formula errors (#N/A / #REF! / #VALUE! / #NAME?) in columns S and T.
    3. Audits column T (Unified Commodity Code) against the 'STANDARD COMM' sheet's column A (valid code list).
       For any row whose T code isn't in that list, tries to resolve it by matching the translated description
       (column S) against 'STANDARD COMM' column B (Commodity Description) and writing back the matching
       column A code. Rows that still can't be resolved are listed at the end for manual review.
    4. Strips leftover AI citation artifacts (e.g. "[cite: 3]") from the 'DESC ITY - ENG' lookup table's
       English-translation column, in case this month's Gemini-drafted rows reintroduced them.

.PARAMETER FilePath
    Path to the TVS.xlsx file. Defaults to the TVS folder inside this project.

.EXAMPLE
    .\Update-TVSTranslations.ps1
    .\Update-TVSTranslations.ps1 -FilePath "C:\Users\User\Desktop\Global Purchase Report Claude\TVS\TVS.xlsx"
#>

param(
    [string]$FilePath = (Join-Path $PSScriptRoot "TVS\TVS.xlsx")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $FilePath)) {
    Write-Host "File not found: $FilePath" -ForegroundColor Red
    exit 1
}

$existingExcel = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($existingExcel) {
    Write-Host "Warning: EXCEL.EXE is already running ($($existingExcel.Count) process(es))." -ForegroundColor Yellow
    Write-Host "Please close all Excel windows first, otherwise this script may fail to open the file cleanly." -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to stop, or wait 5 seconds to continue anyway..."
    Start-Sleep -Seconds 5
}

$excel = $null
$wb = $null

function Release-Com($obj) {
    if ($obj) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null }
}

try {
    Write-Host "Opening $FilePath ..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($FilePath, 0, $false)
    if ($null -eq $wb) { throw "Workbook failed to open (Workbooks.Open returned null)." }
    if ($wb.ReadOnly) { throw "Workbook opened read-only - check if another process/user has it locked." }

    # Main data sheet is always the first sheet, named like "TVS order DDMMYYYY" (changes monthly)
    $ws = $wb.Sheets.Item(1)
    if ($ws.Cells.Item(1,1).Text -ne "Purchase Order Number") {
        throw "Sheet 1 ('$($ws.Name)') doesn't look like the TVS data sheet (expected 'Purchase Order Number' in A1)."
    }
    Write-Host "Data sheet: '$($ws.Name)'"

    $lastRow = $ws.UsedRange.Rows.Count
    Write-Host "Data rows: $($lastRow - 1)"

    # --- Step 1: refresh column S formula down to the last row (adds TRANSLATE() fallback for brand-new terms) ---
    $sFormula = '=IFERROR(VLOOKUP(J2,''DESC ITY - ENG''!B:C,2,FALSE),IFERROR(VLOOKUP(I2,''DESC ITY - ENG''!A:C,3,FALSE),TRANSLATE(J2,"it","en")))'
    $ws.Range("S2").Formula = $sFormula
    $ws.Range("S2").AutoFill($ws.Range("S2:S$lastRow"), 0) | Out-Null
    Write-Host "Column S formula refreshed (rows 2 to $lastRow)."

    # --- Step 2: count formula errors in S and T ---
    # Bulk-checked via SpecialCells (one COM call each) instead of looping .Text over
    # every row individually - a tight per-cell COM loop across ~3000+ rows was found
    # to intermittently throw transient COM errors ("call rejected", "null-valued
    # expression") at a different row each run, i.e. a COM round-trip reliability
    # issue, not a data problem. SpecialCells throws when there are zero matches,
    # which just means the error count is 0.
    $xlCellTypeFormulas = -4123
    $xlErrors = 16
    $errS = 0
    try { $errS = $ws.Range("S2:S$lastRow").SpecialCells($xlCellTypeFormulas, $xlErrors).Count } catch { $errS = 0 }
    $errT = 0
    try { $errT = $ws.Range("T2:T$lastRow").SpecialCells($xlCellTypeFormulas, $xlErrors).Count } catch { $errT = 0 }
    Write-Host "Formula errors - Column S: $errS   Column T: $errT"

    # --- Step 3: audit column T against STANDARD COMM column A, attempt fallback fix via column B match ---
    # Same bulk-read approach: pull each range into a local array with one COM call,
    # do all the matching in memory, and only touch the COM object again to write
    # back the (typically small) subset of cells that actually need fixing.
    $sc = $wb.Sheets.Item("STANDARD COMM")
    $scLastRow = $sc.UsedRange.Rows.Count

    $validCodes = New-Object 'System.Collections.Generic.HashSet[string]'
    $descToCode = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    $scVals = $sc.Range("A2:B$scLastRow").Value2
    for ($i = 1; $i -le $scVals.GetLength(0); $i++) {
        $code = if ($scVals[$i,1] -is [string]) { $scVals[$i,1].Trim() } else { [string]$scVals[$i,1] }
        $descText = if ($scVals[$i,2] -is [string]) { $scVals[$i,2].Trim() } else { "" }
        if ($code -ne "") { [void]$validCodes.Add($code) }
        if ($descText -ne "" -and -not $descToCode.ContainsKey($descText)) { $descToCode[$descText] = $code }
    }

    $stVals = $ws.Range("S2:T$lastRow").Value2
    $mismatchFixed = 0
    $unresolved = @()
    $fixes = @{}  # row (1-based within stVals) -> new code
    for ($i = 1; $i -le $stVals.GetLength(0); $i++) {
        $tVal = if ($stVals[$i,2] -is [string]) { $stVals[$i,2].Trim() } else { "" }
        if ($tVal -eq "" -or $validCodes.Contains($tVal)) { continue }

        $sVal = if ($stVals[$i,1] -is [string]) { $stVals[$i,1].Trim() } else { "" }
        $matchedCode = $null
        foreach ($descKey in $descToCode.Keys) {
            if ($sVal -like "*$descKey*") { $matchedCode = $descToCode[$descKey]; break }
        }
        $r = $i + 1  # actual worksheet row (data starts at row 2)
        if ($matchedCode) {
            $fixes[$r] = $matchedCode
            $mismatchFixed++
        } else {
            $unresolved += "Row $r : T='$tVal' S='$sVal'"
        }
    }
    foreach ($r in $fixes.Keys) { $ws.Cells.Item($r, 20).Value2 = $fixes[$r] }
    Write-Host "T vs STANDARD COMM - auto-fixed: $mismatchFixed   unresolved: $($unresolved.Count)"
    if ($unresolved.Count -gt 0) {
        Write-Host "Unresolved rows (need manual review):" -ForegroundColor Yellow
        $unresolved | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    }

    # --- Step 4: strip leftover AI citation artifacts from the DESC ITY - ENG lookup table ---
    $desc = $wb.Sheets.Item("DESC ITY - ENG")
    $descLastRow = $desc.UsedRange.Rows.Count
    $citePattern = [regex]'\s*\[cite:[^\]]*\]'
    $cleaned = 0
    $descVals = $desc.Range("C2:C$descLastRow").Value2
    $citeFixes = @{}
    for ($i = 1; $i -le $descVals.GetLength(0); $i++) {
        $v = $descVals[$i,1]
        if ($v -is [string] -and $v -match '\[cite') {
            $citeFixes[$i + 1] = $citePattern.Replace($v, '').TrimEnd()
            $cleaned++
        }
    }
    foreach ($r in $citeFixes.Keys) { $desc.Cells.Item($r, 3).Value2 = $citeFixes[$r] }
    Write-Host "Citation artifacts cleaned in 'DESC ITY - ENG': $cleaned"

    # Re-refresh column S once more so cleaned lookup text/new fixes are reflected in this month's rows
    $ws.Range("S2").Formula = $sFormula
    $ws.Range("S2").AutoFill($ws.Range("S2:S$lastRow"), 0) | Out-Null

    # This AutoFill can re-trigger TRANSLATE()'s asynchronous (web-service-backed)
    # recalculation for any newly-changed lookups, so give it a moment before saving.
    $excel.CalculateUntilAsyncQueriesDone()
    for ($i = 0; $i -lt 20 -and $excel.CalculationState -ne 0; $i++) { Start-Sleep -Milliseconds 250 }

    $beforeSaveTime = Get-Date
    $wb.Save()
    # Workbook.Saved is an unreliable read immediately after Save() over COM (found
    # by testing: the file's actual on-disk content/timestamp were already correct
    # even when .Saved falsely reported $false) - verify against the file itself
    # instead of trusting that in-memory property.
    Start-Sleep -Milliseconds 500
    $onDiskTime = (Get-Item $FilePath).LastWriteTime
    if ($onDiskTime -lt $beforeSaveTime) { throw "Save() completed but the file's on-disk timestamp didn't update - it may not have actually saved." }
    Write-Host "Saved successfully (confirmed via on-disk file timestamp)." -ForegroundColor Green

    Write-Host ""
    Write-Host "=== Summary ===" -ForegroundColor Cyan
    Write-Host "Rows processed        : $($lastRow - 1)"
    Write-Host "Formula errors (S/T)  : $errS / $errT"
    Write-Host "T codes auto-fixed    : $mismatchFixed"
    Write-Host "T codes unresolved    : $($unresolved.Count)"
    Write-Host "Citation text cleaned : $cleaned"
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}
finally {
    if ($wb) { $wb.Close($false); Release-Com $wb }
    if ($excel) { $excel.Quit(); Release-Com $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $stillRunning = $null
    for ($i = 0; $i -lt 6; $i++) {
        Start-Sleep -Seconds 1
        $stillRunning = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
        if (-not $stillRunning) { break }
    }
    if ($stillRunning) {
        Write-Host "Warning: EXCEL.EXE still running after this script ($($stillRunning.Count) process(es))." -ForegroundColor Yellow
        Write-Host "If this persists, close them manually via Task Manager before running this script again." -ForegroundColor Yellow
    }
}
