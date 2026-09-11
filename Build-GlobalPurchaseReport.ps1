<#
.SYNOPSIS
    Monthly Global Purchase Report builder (Script A / Excel COM) - replaces the old UiPath RPA
    "merge + report generation" steps. All 7 entities supported (MAZAL added once its June 2026 data
    arrived); MAZAL is skipped automatically with a warning if MAZ\MAZ.xlsx isn't present yet.

.DESCRIPTION
    Mechanically replicates the exact copy/paste + formula-fill operations found in the old RPA's
    Main.xaml, per entity, so the resulting Report is structurally identical to what the RPA produced
    (MAZAL has no RPA automation to replicate - its mapping was instead reverse-engineered empirically,
    same method as TVS: cross-referencing known-correct historical Report rows against the raw file):

      MHK / STC : raw ERP ACT001 export -> pasted into that entity's own
                  "Trimmed Program\<Entity> Act001 Trimmed.xlsx" (sheet 'ACT001', column-shifted per
                  entity) -> sheet 'Trimmed' recomputes via its own existing formulas (row 3 autofilled
                  down) -> Trimmed!A:T appended into Report.
      MZQ / MIL : no Trimmed file. Cumulative vendor file filtered to the target month (date column B)
                  -> columns A:P and T copied directly into Report (Q:S filled by Report's own formulas).
      RACO / TVS: cumulative vendor/finance file filtered to the target month (RACO: column B date,
                  TVS: column G Invoice Date) -> pasted directly into that entity's
                  "Trimmed Program\<Entity> Act001 Trimmed '<MM-yyyy>.xlsx" sheet 'Trimmed' (A:T) ->
                  appended into Report.
      MAZAL     : no Trimmed file. Cumulative vendor file (sheet 'MHK' - legacy misnamed, like MZQ)
                  filtered to the target month (Purchase Order Date, column B) -> columns A:J copied
                  directly, PBRT_TYPE (K) inserted as a constant "B", columns L:P copied directly (NO
                  shift - raw's unused Division column and Report's inserted PBRT_TYPE constant sit at
                  the same position and cancel out) -> appended into Report with DB constant "MAZ".

    After all entities are appended: the Exchange Rate sheet is updated first (see
    Update-ExchangeRateSheet), then Report's Q:S (exchange rate / USD price) and the AA/AB-independent
    part of U:AJ formulas (U,V,W,Q,R,S,Y,Z,AJ) are filled for the new rows only, ONE COLUMN AT A TIME
    with an immediate freeze-to-values after each column (see Set-FormulaColumnsFrozen) - filling and
    calculating all formula columns together would recalculate the full dependency web across every new
    row at once, which is what caused Excel to hang for 10+ minutes earlier in this project. Historical
    rows from prior months stay frozen to static values and are never re-touched. X (period index) is
    deliberately left out of this entirely - it's kept as a permanently-live formula pre-filled far ahead
    (X2:X70000) in the Temp reference file, since it's cheap (no full-column scan) and Q depends on it.

    AA/AB (Max/Min PO Price) are deliberately NOT filled by this script (per explicit instruction,
    2026-08): the report owner copies AA2:AB2's live formula down to the new rows manually in Excel and
    confirms it's done before continuing. This script stops after the first-pass columns above and prints
    the exact next steps. Once AA/AB are filled manually, run Complete-FormulaColumns.ps1 to fill the
    remaining AA/AB-dependent columns (AC/AD/AE/AG/AH, plus AF/AI which are grouped with them) and the
    New Vendors review sheet - see that script's header for why it's a separate step and how it finds the
    row range without relying on any formula column being filled yet.

    Email send and MMS upload are NOT part of this script (stay manual, same as the MTC project).

    PERFORMANCE / RELIABILITY NOTE: all Excel reads/writes use bulk Range.Value2 array operations
    (whole rectangular blocks in one COM call), never per-cell loops. An earlier per-cell version of
    this script locked up Excel (COM went unresponsive - "RPC server is unavailable") when processing
    MHK's ~800 rows one cell at a time; bulk range I/O avoids that entirely and is far faster.

.PARAMETER WorkingFolder
    Folder containing this month's per-entity source files (MHK\MHK.xlsx, STC\STC.xlsx, MZQ\MZQ.xlsx,
    MIL\MIL.xlsx, RACO\RACO.xlsx, TVS\TVS.xlsx, MAZ\MAZ.xlsx). Defaults to this script's folder. MAZ\MAZ.xlsx
    is optional - if missing, MAZAL is skipped with a warning rather than failing the whole run.

.PARAMETER TemplateFolder
    Folder containing all four entities' undated "<Entity> Act001 Trimmed.xlsx" template files. RACO
    and TVS always read this same fixed template each run (matching the original RPA's own pattern),
    then Save As a new dated snapshot into WorkingFolder\RACO|TVS - never overwriting the template
    itself, so there's no source/destination filename collision.

.PARAMETER ReportTemplatePath
    The existing master Report workbook to append this month's rows into (a copy is made - the
    original template is never modified in place). If omitted, auto-detects the latest prior monthly
    output in WorkingFolder\REPORT (so each month builds on last month's, not a static reference file);
    falls back to the RPA reference "-Temp.xlsx" only if no prior output exists yet.

.PARAMETER RacoTrimmedFile
    Full path to the RACO Trimmed template. Defaults to the fixed "RACO Act001 Trimmed.xlsx" in
    TemplateFolder (read-only source; its data is overwritten in memory each run, then Save As'd to a
    new dated snapshot - the template file itself is never modified).

.PARAMETER TvsTrimmedFile
    Same as -RacoTrimmedFile, but for TVS.

.PARAMETER ReportMonth
    Target month to filter cumulative source files down to, format "yyyy-MM" (e.g. "2026-06").

.PARAMETER OutputPath
    Where to save the new dated Report. Defaults to WorkingFolder\REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_<ReportMonth>.xlsx

.PARAMETER ExchangeRatePath
    Path to the monthly "STANDARD EXCHANGE RATE" file (from the manual Gmail step - see SOP 2.1). If
    omitted, looks for "STANDARD EXCHANGE RATE.xls" then ".xlsx" in WorkingFolder\Exchange Rate\. If not
    found, exchange rate update is skipped with a warning (doesn't fail the run - the report's Q:S
    columns will just stay blank until the rate is added and the file is refreshed in Excel).

.EXAMPLE
    .\Build-GlobalPurchaseReport.ps1 -ReportMonth "2026-06"
#>

param(
    [string]$WorkingFolder = $PSScriptRoot,
    [string]$TemplateFolder = (Join-Path $PSScriptRoot "Global Purchase Report RPA\Trimmed Program"),
    [string]$ReportTemplatePath,   # auto-detected below if not given: latest prior monthly output, else the RPA reference file
    [string]$RacoTrimmedFile,      # auto-detected below if not given: latest "RACO Act001 Trimmed*.xlsx" across WorkingFolder\RACO and TemplateFolder
    [string]$TvsTrimmedFile,       # auto-detected below if not given: latest "TVS Act001 Trimmed*.xlsx" across WorkingFolder\TVS and TemplateFolder
    [string]$ExchangeRatePath,     # auto-detected below if not given: WorkingFolder\Exchange Rate\STANDARD EXCHANGE RATE.xls|xlsx
    [Parameter(Mandatory = $true)]
    [string]$ReportMonth,
    [string]$OutputPath,
    [switch]$Force   # skip the duplicate-month safety check
)

$ErrorActionPreference = "Stop"

if (-not $OutputPath) {
    $OutputPath = Join-Path $WorkingFolder "REPORT\FIN_MHK+STC+MZQ+MIL+RACO+TVS_PUR_REPORT_$ReportMonth.xlsx"
}
$reportOutDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $reportOutDir)) { New-Item -ItemType Directory -Path $reportOutDir -Force | Out-Null }

$monthDate = [datetime]::ParseExact($ReportMonth, "yyyy-MM", $null)
$monthSuffix = $monthDate.ToString("MM-yyyy")   # matches old RPA naming, e.g. "06-2026"

# --- Auto-detect continuity inputs, so this script doesn't need editing every month ---
# (Found this the hard way: hardcoded dated filenames as parameter defaults go stale after one run.)

if (-not $ReportTemplatePath) {
    $priorOutputs = Get-ChildItem (Join-Path $WorkingFolder "REPORT") -Filter "FIN_*_PUR_REPORT_*.xlsx" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending
    if ($priorOutputs) {
        $ReportTemplatePath = $priorOutputs[0].FullName
        Write-Host "Auto-detected ReportTemplatePath (latest prior monthly output): $ReportTemplatePath"
    } else {
        $ReportTemplatePath = Join-Path $PSScriptRoot "Global Purchase Report RPA\REPORT\FIN_MHK+STC+MZQ+MIL+RACO_PUR_REPORT-Temp.xlsx"
        Write-Host "Auto-detected ReportTemplatePath (no prior monthly output found, using RPA reference file): $ReportTemplatePath"
    }
}

if (-not $RacoTrimmedFile) {
    # Always the fixed undated template (never a dated snapshot) - matches the original RPA's own
    # pattern (open this same file every month, then Save As a new dated snapshot). Using "latest
    # dated snapshot" here instead caused a same-filename collision when re-running the same month
    # (its own prior snapshot got picked as the read-only source AND the SaveAs target).
    $RacoTrimmedFile = Join-Path $TemplateFolder "RACO Act001 Trimmed.xlsx"
    if (-not (Test-Path $RacoTrimmedFile)) { throw "'$RacoTrimmedFile' not found - pass -RacoTrimmedFile explicitly." }
    Write-Host "RacoTrimmedFile (fixed template): $RacoTrimmedFile"
}

if (-not $TvsTrimmedFile) {
    $TvsTrimmedFile = Join-Path $TemplateFolder "TVS Act001 Trimmed.xlsx"
    if (-not (Test-Path $TvsTrimmedFile)) { throw "'$TvsTrimmedFile' not found - pass -TvsTrimmedFile explicitly." }
    Write-Host "TvsTrimmedFile (fixed template): $TvsTrimmedFile"
}

if (-not $ExchangeRatePath) {
    $erFolder = Join-Path $WorkingFolder "Exchange Rate"
    $erCandidate = @("STANDARD EXCHANGE RATE.xls", "STANDARD EXCHANGE RATE.xlsx") |
        ForEach-Object { Join-Path $erFolder $_ } | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($erCandidate) {
        $ExchangeRatePath = $erCandidate
        Write-Host "Auto-detected ExchangeRatePath: $ExchangeRatePath"
    } else {
        Write-Host "No 'STANDARD EXCHANGE RATE.xls[x]' found in '$erFolder' - exchange rate update will be skipped (Q:S columns stay blank until it's added)." -ForegroundColor Yellow
    }
}

function Release-Com($obj) { if ($obj) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null } }

function New-ExcelApp {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    return $excel
}

function Open-Workbook($excel, $path, [bool]$readOnly = $false) {
    $wb = $excel.Workbooks.Open($path, 0, $readOnly)
    if ($null -eq $wb) { throw "Failed to open: $path" }
    if (-not $readOnly -and $wb.ReadOnly) { throw "Opened read-only unexpectedly (file locked?): $path" }
    return $wb
}

function Get-UsedLastRow($ws) { return $ws.UsedRange.Rows.Count }

# UsedRange can be padded with blank formatted rows (found a ~4,649-row gap in the real production
# Report file this way) - find the TRUE last row with data in a given column instead. Matches what the
# old RPA's own "Find First/Last Data Row" activities were presumably doing, rather than trusting UsedRange.
function Get-TrueLastDataRow($ws, [int]$col) {
    $lastCell = $ws.Cells.Item($ws.Rows.Count, $col).End(-4162)   # xlUp
    return $lastCell.Row
}

# Bulk-reads a rectangular block in ONE COM call and returns a normalized 0-based object[,] array.
function Get-RangeArray($ws, [int]$r1, [int]$c1, [int]$r2, [int]$c2) {
    # NOTE: every return below uses ",$x" (comma operator) - without it, PowerShell enumerates/flattens
    # the 2D array across the function-return boundary into a plain 1D array, silently corrupting the
    # shape (this bit us hard during testing: GetLength(1) failing with "index outside bounds" downstream).
    if ($r2 -lt $r1 -or $c2 -lt $c1) { return , (New-Object 'object[,]' 0, 0) }
    $rows = $r2 - $r1 + 1
    $cols = $c2 - $c1 + 1
    $range = $ws.Range($ws.Cells.Item($r1, $c1), $ws.Cells.Item($r2, $c2))
    $v = $range.Value2
    $arr = New-Object 'object[,]' $rows, $cols
    if ($rows -eq 1 -and $cols -eq 1) {
        $arr[0, 0] = $v
        return , $arr
    }
    for ($i = 1; $i -le $rows; $i++) {
        for ($j = 1; $j -le $cols; $j++) { $arr[($i - 1), ($j - 1)] = $v[$i, $j] }
    }
    return , $arr
}

# Bulk-writes a 0-based object[,] array starting at (r1,c1) in ONE COM call.
function Set-RangeArray($ws, [int]$r1, [int]$c1, $arr) {
    $rows = $arr.GetLength(0)
    $cols = $arr.GetLength(1)
    if ($rows -eq 0 -or $cols -eq 0) { return }
    $range = $ws.Range($ws.Cells.Item($r1, $c1), $ws.Cells.Item($r1 + $rows - 1, $c1 + $cols - 1))
    $range.Value2 = $arr
}

# Extracts a column sub-range [c1..c2] (0-based, inclusive) from a 0-based 2D array - pure in-memory, no COM.
function Slice-Cols($arr, [int]$c1, [int]$c2) {
    $rows = $arr.GetLength(0)
    $cols = $c2 - $c1 + 1
    $result = New-Object 'object[,]' $rows, $cols
    for ($i = 0; $i -lt $rows; $i++) { for ($j = 0; $j -lt $cols; $j++) { $result[$i, $j] = $arr[$i, ($c1 + $j)] } }
    return , $result
}

function Parse-CellDate($v) {
    if ($v -is [double]) { return ([datetime]'1899-12-30').AddDays($v) }
    if ($v -is [datetime]) { return $v }
    if ($v -is [string] -and $v -ne "") { try { return [datetime]::Parse($v) } catch { return $null } }
    return $null
}

# Given a full bulk-read array (0-based) and a 0-based date column index within it, returns the row
# offsets (0-based, into $arr) whose date falls in $month ("yyyy-MM"). Pure in-memory - no COM calls.
function Get-CurrentMonthRowOffsets($arr, [int]$dateColIndex, [string]$month) {
    $target = [datetime]::ParseExact($month, "yyyy-MM", $null)
    $offsets = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $arr.GetLength(0); $i++) {
        $d = Parse-CellDate $arr[$i, $dateColIndex]
        if ($d -and $d.Year -eq $target.Year -and $d.Month -eq $target.Month) { [void]$offsets.Add($i) }
    }
    return , $offsets
}

function Select-RowOffsets($arr, $offsets) {
    $cols = $arr.GetLength(1)
    $result = New-Object 'object[,]' $offsets.Count, $cols
    for ($i = 0; $i -lt $offsets.Count; $i++) {
        for ($c = 0; $c -lt $cols; $c++) { $result[$i, $c] = $arr[$offsets[$i], $c] }
    }
    return , $result
}

# --- Per-entity processors. Each returns a 2D object[,] of rows shaped A:T (20 cols) ready for Report. ---

function Process-MHK($excel, $sourcePath) {
    Write-Host "[MHK] Processing $sourcePath"
    $srcWb = Open-Workbook $excel $sourcePath $true
    $raw = $null
    try {
        $srcWs = $srcWb.Sheets.Item(1)
        $lastRow = Get-UsedLastRow $srcWs
        Write-Host "[MHK] Source rows: $($lastRow - 1) (single-month export, no filter needed)"
        if ($lastRow -ge 2) { $raw = Get-RangeArray $srcWs 2 1 $lastRow 27 } else { $raw = New-Object 'object[,]' 0, 27 }
    } finally { $srcWb.Close($false) | Out-Null; Release-Com $srcWb }

    $tplSrc = Join-Path $TemplateFolder "MHK Act001 Trimmed.xlsx"
    $tplPath = Join-Path $WorkingFolder "MHK\MHK Act001 Trimmed.xlsx"
    Copy-Item $tplSrc $tplPath -Force   # work on a copy - never write into the RPA reference folder
    $wb = Open-Workbook $excel $tplPath $false
    try {
        $act = $wb.Sheets.Item("ACT001")
        $trimmed = $wb.Sheets.Item("Trimmed")
        $act.Range("A3:W5000").ClearContents() | Out-Null
        # NOTE: only column A gets cleared, matching the old RPA's exact "a3:A5000" clear range - Trimmed!B:T
        # are LIVE FORMULAS in row 3 that AutoFill propagates down; clearing the whole A:T range destroys
        # that master formula row and leaves every appended row blank (found this the hard way in testing).
        $trimmed.Range("A3:A5000").ClearContents() | Out-Null

        $n = $raw.GetLength(0)
        if ($n -gt 0) {
            # Build the ACT001 block (23 cols, A:W) in-memory from raw (27 cols) per the extracted mapping:
            # A:E<-B:F  F<-H  G:J<-L:O  K<-P  M:W<-Q:AA   (0-based raw indices: B=1..F=5, H=7, L=11..O=14, P=15, Q=16..AA=26)
            $actArr = New-Object 'object[,]' $n, 23
            for ($i = 0; $i -lt $n; $i++) {
                for ($k = 0; $k -lt 5; $k++) { $actArr[$i, $k] = $raw[$i, (1 + $k)] }          # A:E <- B:F
                $actArr[$i, 5] = $raw[$i, 7]                                                    # F   <- H
                for ($k = 0; $k -lt 4; $k++) { $actArr[$i, (6 + $k)] = $raw[$i, (11 + $k)] }    # G:J <- L:O
                $actArr[$i, 10] = $raw[$i, 15]                                                  # K   <- P
                for ($k = 0; $k -lt 11; $k++) { $actArr[$i, (12 + $k)] = $raw[$i, (16 + $k)] }  # M:W <- Q:AA
            }
            Set-RangeArray $act 3 1 $actArr | Out-Null

            $lastR = 3 + $n - 1
            $anchorA = Slice-Cols $actArr 4 4   # Trimmed!A <- ACT001!E (col index 4)
            Set-RangeArray $trimmed 3 1 $anchorA | Out-Null

            $trimmed.Range("B3:T3").AutoFill($trimmed.Range("B3:T$lastR"), 0) | Out-Null

            # Trimmed!F/G (VEND_INV_NO/REFERENCE) are NOT formulas in the template - the old RPA never
            # carried these two raw columns into ACT001 at all (a real gap). Patch directly from raw.
            $fg = New-Object 'object[,]' $n, 2
            for ($i = 0; $i -lt $n; $i++) { $fg[$i, 0] = $raw[$i, 8]; $fg[$i, 1] = $raw[$i, 11] }  # F<-raw col9, G<-raw col12
            Set-RangeArray $trimmed 3 6 $fg | Out-Null

            # Trimmed!B:T are live formulas referencing 'ACT001' - COM Value2 writes don't reliably mark
            # them dirty for recalculation, so force it before reading back (else Get-RangeArray reads stale/blank cache).
            $trimmed.Calculate() | Out-Null
            $result = Get-RangeArray $trimmed 3 1 $lastR 20
        } else {
            $result = New-Object 'object[,]' 0, 20
        }
        $wb.Save() | Out-Null
        return , $result
    } finally { $wb.Close($false) | Out-Null; Release-Com $wb }
}

function Process-STC($excel, $sourcePath) {
    Write-Host "[STC] Processing $sourcePath"
    $srcWb = Open-Workbook $excel $sourcePath $true
    $raw = $null
    try {
        $srcWs = $srcWb.Sheets.Item(1)
        $lastRow = Get-UsedLastRow $srcWs
        Write-Host "[STC] Source rows: $($lastRow - 1) (single-month export, no filter needed)"
        if ($lastRow -ge 2) { $raw = Get-RangeArray $srcWs 2 1 $lastRow 23 } else { $raw = New-Object 'object[,]' 0, 23 }
    } finally { $srcWb.Close($false) | Out-Null; Release-Com $srcWb }

    $tplSrc = Join-Path $TemplateFolder "STC Act001 Trimmed.xlsx"
    $tplPath = Join-Path $WorkingFolder "STC\STC Act001 Trimmed.xlsx"
    Copy-Item $tplSrc $tplPath -Force
    $wb = Open-Workbook $excel $tplPath $false
    try {
        $act = $wb.Sheets.Item("ACT001")
        $trimmed = $wb.Sheets.Item("Trimmed")
        $act.Range("A3:W5000").ClearContents() | Out-Null
        # NOTE: only column A gets cleared, matching the old RPA's exact "a3:A5000" clear range - Trimmed!B:T
        # are LIVE FORMULAS in row 3 that AutoFill propagates down; clearing the whole A:T range destroys
        # that master formula row and leaves every appended row blank (found this the hard way in testing).
        $trimmed.Range("A3:A5000").ClearContents() | Out-Null

        $n = $raw.GetLength(0)
        if ($n -gt 0) {
            Set-RangeArray $act 3 1 $raw | Out-Null   # A:W <- A:W direct 1:1

            $lastR = 3 + $n - 1
            $anchorA = Slice-Cols $raw 5 5   # Trimmed!A <- ACT001!F (col index 5)
            $anchorC = Slice-Cols $raw 7 7   # Trimmed!C <- ACT001!H (col index 7)
            Set-RangeArray $trimmed 3 1 $anchorA | Out-Null
            Set-RangeArray $trimmed 3 3 $anchorC | Out-Null

            # Column C is a raw-copied value (not a formula) - autofill B and D:T separately so C isn't overwritten
            $trimmed.Range("B3").AutoFill($trimmed.Range("B3:B$lastR"), 0) | Out-Null
            $trimmed.Range("D3:T3").AutoFill($trimmed.Range("D3:T$lastR"), 0) | Out-Null

            # Force recalculation before reading back - see MHK's comment above for why.
            $trimmed.Calculate() | Out-Null
            $result = Get-RangeArray $trimmed 3 1 $lastR 20
        } else {
            $result = New-Object 'object[,]' 0, 20
        }
        $wb.Save() | Out-Null
        return , $result
    } finally { $wb.Close($false) | Out-Null; Release-Com $wb }
}

function Process-CumulativeDirect($excel, $sourcePath, [string]$sheetName, [int]$dateColIndex1Based, [string]$month, [string]$label) {
    Write-Host "[$label] Processing $sourcePath"
    $wb = Open-Workbook $excel $sourcePath $true
    try {
        $ws = $wb.Sheets.Item($sheetName)
        $lastRow = Get-UsedLastRow $ws
        if ($lastRow -lt 2) { return , (New-Object 'object[,]' 0, 20) }
        $full = Get-RangeArray $ws 2 1 $lastRow 20   # A:T
        $offsets = Get-CurrentMonthRowOffsets $full ($dateColIndex1Based - 1) $month
        Write-Host "[$label] $($offsets.Count) rows matched month $month (of $($lastRow-1) total)"
        $matched = Select-RowOffsets $full $offsets

        if ($label -in @("MZQ", "MIL")) {
            # Report only takes A:P and T from these two - zero out Q:S/U-ish columns (16,17,18) so they
            # don't carry stray source data into positions Report will overwrite via formula anyway.
            for ($i = 0; $i -lt $matched.GetLength(0); $i++) {
                for ($c = 16; $c -le 18; $c++) { $matched[$i, $c] = $null }
            }
        }

        if ($label -eq "MIL") {
            # MIL's raw column K (PBRT_TYPE position) is actually a "Div" division code (1,2,6,7,8,9,
            # sometimes "M"), not a B/O part-type flag - MZQ's own column K is genuinely B/O already, so
            # this is MIL-specific. Found because it silently broke AA-AG's "which vendor hit the max
            # price" VLOOKUP (the B/O value feeds into the AC/AD lookup key, so a stray division number
            # there means it never matches any "B"-keyed row). Normalize any numeric 1-9 value to "B".
            for ($i = 0; $i -lt $matched.GetLength(0); $i++) {
                $v = $matched[$i, 10]
                if ($v -is [double] -and $v -ge 1 -and $v -le 9) { $matched[$i, 10] = "B" }
                elseif ($v -is [string] -and $v -match '^[1-9]$') { $matched[$i, 10] = "B" }
            }
        }
        return , $matched
    } finally { $wb.Close($false) | Out-Null; Release-Com $wb }
}

function Process-RACO($excel, $sourcePath, [string]$month) {
    $wb = Open-Workbook $excel $sourcePath $true
    $sheetName = "RACO收货记录"
    $sheetExists = $false
    foreach ($s in $wb.Sheets) { if ($s.Name -eq $sheetName) { $sheetExists = $true } }
    if (-not $sheetExists) {
        Write-Host "[RACO] Sheet '$sheetName' not found, falling back to first sheet" -ForegroundColor Yellow
        $sheetName = $wb.Sheets.Item(1).Name
    }
    $wb.Close($false) | Out-Null; Release-Com $wb

    $rowsArr = Process-CumulativeDirect $excel $sourcePath $sheetName 2 $month "RACO"

    $tplPath = $RacoTrimmedFile   # already a full path (auto-detected above, or passed explicitly)
    $twb = Open-Workbook $excel $tplPath $true   # read-only: template folder's snapshots are historical, never overwritten
    try {
        $trimmed = $twb.Sheets.Item("Trimmed")
        $trimmed.Range("A2:T5000").ClearContents() | Out-Null
        if ($rowsArr.GetLength(0) -gt 0) { Set-RangeArray $trimmed 2 1 $rowsArr | Out-Null }
        $outSnapshot = Join-Path $WorkingFolder "RACO\RACO Act001 Trimmed '$monthSuffix.xlsx"
        $twb.SaveAs($outSnapshot) | Out-Null
        Write-Host "[RACO] New month's snapshot saved to $outSnapshot (template folder untouched)"
        return , $rowsArr
    } finally { $twb.Close($false) | Out-Null; Release-Com $twb }
}

function Process-TVS($excel, $sourcePath, [string]$month) {
    Write-Host "[TVS] Processing $sourcePath"
    $wb = Open-Workbook $excel $sourcePath $true
    $rowsArr = $null
    try {
        $ws = $wb.Sheets.Item(1)   # "TVS order DDMMYYYY" - name changes monthly, always sheet 1
        if ($ws.Cells.Item(1, 1).Text -ne "Purchase Order Number") {
            throw "[TVS] Sheet 1 ('$($ws.Name)') doesn't look like the TVS data sheet."
        }
        $lastRow = Get-UsedLastRow $ws
        if ($lastRow -lt 2) {
            $rowsArr = New-Object 'object[,]' 0, 20
        } else {
            # Raw TVS columns (22, confirmed by cross-referencing real historical Report rows against
            # this file - the raw layout does NOT match Report's A:T positionally, unlike RACO):
            #   0 PO Number, 1 PO Date, 2 PO issue date, 3 Vendor ID, 4 VendorName, 5 Vendor_Invoice_No,
            #   6 Invoice Date, 7 REF_NO, 8 Item Code, 9 Part Description (Italian), 10 Commodity(IT),
            #   11 Commodity(EN), 12 Division, 13 Quantity, 14 Unit, 15 PO_Price, 16 PO_Unit_Price,
            #   17 Total_PO_Price, 18 Translated Part Description(EN), 19 Unified Commodity Code
            $full = Get-RangeArray $ws 2 1 $lastRow 20
            $offsets = Get-CurrentMonthRowOffsets $full 6 $month   # column G (index 6) = Invoice Date
            Write-Host "[TVS] $($offsets.Count) rows matched month $month (of $($lastRow-1) total)"
            $src = Select-RowOffsets $full $offsets

            # Confirmed empirically: found existing TVS rows in the reference Report (filtered col T=
            # "TVS"), matched their PO numbers back to this raw file, and compared field-by-field.
            $n = $src.GetLength(0)
            $rowsArr = New-Object 'object[,]' $n, 20
            for ($i = 0; $i -lt $n; $i++) {
                $rowsArr[$i, 0] = $src[$i, 0]                                            # PO_NO <- PO Number
                $rowsArr[$i, 1] = $src[$i, 6]                                            # PO_REC_DATE <- Invoice Date
                $rowsArr[$i, 2] = $src[$i, 1]                                            # PO_ISS_DATE <- PO Date
                $rowsArr[$i, 3] = $src[$i, 3]                                            # VENDOR_ID <- Vendor ID
                $rowsArr[$i, 4] = $src[$i, 4]                                            # VENDOR_NAME <- VendorName
                $rowsArr[$i, 5] = $src[$i, 5]                                            # VEND_INV_NO <- Vendor_Invoice_No
                $rowsArr[$i, 6] = $src[$i, 7]                                            # REFERENCE <- REF_NO
                $rowsArr[$i, 7] = $src[$i, 8]                                            # ITEM_CODE <- Item Code
                $rowsArr[$i, 8] = $src[$i, 9]                                            # PART_DESC <- Part Description (Italian)
                $rowsArr[$i, 9] = $src[$i, 19]                                           # COMM_CODE <- Unified Commodity Code
                $rowsArr[$i, 10] = "B"                                                   # PBRT_TYPE <- constant (confirmed always "B" in historical rows)
                $rowsArr[$i, 11] = $src[$i, 13]                                          # REC_QTY <- Quantity
                $rowsArr[$i, 12] = $src[$i, 14]                                          # UOM <- Unit
                $rowsArr[$i, 13] = "EUR"                                                 # PO_CUR <- constant (Italian vendors, confirmed always EUR)
                $qty = $src[$i, 13]; $tot = $src[$i, 17]
                $rowsArr[$i, 14] = if ($qty -and $qty -ne 0) { $tot / $qty } else { $null }  # PO_PRICE_FCY <- Total/Qty (raw's own unit-price col didn't match observed values)
                $rowsArr[$i, 15] = $tot                                                  # TOT_PO_PRICE_FCY <- Total_PO_Price_PO CUR
                # 16:18 (EXCH_RATE/PO_PRICE_USD/TOT_PO_PRICE_USD) left null - filled by Report's own Q:S formulas
                $rowsArr[$i, 19] = "TVS"                                                 # DB <- constant
            }
        }
    } finally { $wb.Close($false) | Out-Null; Release-Com $wb }

    $tplPath = $TvsTrimmedFile   # already a full path (auto-detected above, or passed explicitly)
    $twb = Open-Workbook $excel $tplPath $true
    try {
        $trimmed = $twb.Sheets.Item("Trimmed")
        $trimmed.Range("A2:A5000").ClearContents() | Out-Null
        if ($rowsArr.GetLength(0) -gt 0) { Set-RangeArray $trimmed 2 1 $rowsArr | Out-Null }
        $outSnapshot = Join-Path $WorkingFolder "TVS\TVS Act001 Trimmed '$monthSuffix.xlsx"
        $twb.SaveAs($outSnapshot) | Out-Null
        Write-Host "[TVS] New month's snapshot saved to $outSnapshot (template folder untouched)"
        return , $rowsArr
    } finally { $twb.Close($false) | Out-Null; Release-Com $twb }
}

function Process-MAZAL($excel, $sourcePath, [string]$month) {
    Write-Host "[MAZAL] Processing $sourcePath"
    $wb = Open-Workbook $excel $sourcePath $true
    $rowsArr = $null
    try {
        $ws = $wb.Sheets.Item("MHK")   # legacy in-file sheet name (like MZQ) - this IS MAZAL's real data
        $lastRow = Get-UsedLastRow $ws
        if ($lastRow -lt 2) {
            $rowsArr = New-Object 'object[,]' 0, 20
        } else {
            # Raw MAZAL columns (17, confirmed by cross-referencing real historical Report rows - "OC-13367"/
            # "OC-13500" - against this file: a clean 1:1 mapping, no reshuffle, unlike TVS):
            #   0 PO Number, 1 PO Date, 2 PO issue date, 3 Vendor ID, 4 VendorName, 5 Vendor_Invoice_No,
            #   6 REF_NO, 7 Item Code, 8 Part Description, 9 Commodity, 10 Division (unused - PBRT_TYPE is
            #   a constant "B", confirmed - Division holds "1" in every sampled row, doesn't match "B"),
            #   11 Quantity, 12 Unit, 13 PO_Price_PO CUR (actually holds the currency code, e.g. "USD"),
            #   14 PO UNIT PRICE PO_CUR, 15 Total_PO_Price_PO CUR
            $full = Get-RangeArray $ws 2 1 $lastRow 16
            $offsets = Get-CurrentMonthRowOffsets $full 1 $month   # column B (index 1) = Purchase Order Date
            Write-Host "[MAZAL] $($offsets.Count) rows matched month $month (of $($lastRow-1) total)"
            $src = Select-RowOffsets $full $offsets

            $n = $src.GetLength(0)
            $rowsArr = New-Object 'object[,]' $n, 20
            for ($i = 0; $i -lt $n; $i++) {
                for ($c = 0; $c -lt 10; $c++) { $rowsArr[$i, $c] = $src[$i, $c] }             # A:J <- raw 0:9 (direct)
                $rowsArr[$i, 10] = "B"                                                        # PBRT_TYPE <- constant
                # L:P <- raw 11:15 (direct, NO shift): raw's unused Division sits at index 10, same position
                # as Report's inserted PBRT_TYPE constant - the two insertions/drops cancel out positionally,
                # so columns after that line up 1:1. (Bug found in production: originally read raw[c-1],
                # which pulled raw K:O - Division through PO-unit-price - into Report's L:P instead of raw L:P.)
                for ($c = 11; $c -lt 16; $c++) { $rowsArr[$i, $c] = $src[$i, $c] }
                # 16:18 (EXCH_RATE/PO_PRICE_USD/TOT_PO_PRICE_USD) left null - filled by Report's own Q:S formulas
                $rowsArr[$i, 19] = "MAZ"                                                      # DB <- constant (confirmed via historical rows, not "MAZAL")
            }
        }
        return , $rowsArr
    } finally { $wb.Close($false) | Out-Null; Release-Com $wb }
}

# Reads the monthly "STANDARD EXCHANGE RATE" file (HK-dollar-denominated rates) and writes them
# DIRECTLY into the Report's "Exchange Rate" sheet Block 1 (cols B:M, "...=> HKD", no conversion) for
# $month. Block 2 (cols O:Z, "...=> USD") is NOT computed here - it's a live formula block already in
# the sheet (e.g. P{row} = "=IF(C{row}="",...,ROUND(+C{row}/K{row},4))", referencing Block 1's own row),
# so once Block 1 has values, Block 2 computes itself. This function only needs to make sure Block 2's
# formula pattern is present for the target row (copied from an adjacent row if the row is newly created).
#
# CORRECTION (found in review): an earlier version of this function computed the USD conversion itself
# and wrote static numbers into Block 2 - that overwrote the sheet's own live formulas with hardcoded
# values and, worse, never touched Block 1 at all (so Block 1 stayed blank every month). Fixed to match
# the sheet's actual two-block design.
function Update-ExchangeRateSheet($excel, [string]$ratePath, $reportWs, [string]$month) {
    Write-Host "[ExchangeRate] Processing $ratePath"
    $rwb = Open-Workbook $excel $ratePath $true
    $hkdRates = @{}   # currency code -> "1 unit foreign currency = X HKD", straight from source, no math
    try {
        $rws = $rwb.Sheets.Item("Exchange rate table")
        if ($rws.Cells.Item(12, 1).Text.Trim() -ne "CURRENCY" -or $rws.Cells.Item(12, 5).Text.Trim() -ne "HK DOLLAR") {
            throw "[ExchangeRate] Header row (12) doesn't match the expected layout (A12='CURRENCY', E12='HK DOLLAR') - file structure may have changed."
        }
        $r = 13
        while ($rws.Cells.Item($r, 3).Text.Trim() -ne "") {
            $code = $rws.Cells.Item($r, 3).Text.Trim()
            $rate = $rws.Cells.Item($r, 5).Value2
            if ($code -and $rate -is [double]) { $hkdRates[$code] = $rate }
            $r++
        }
    } finally { $rwb.Close($false) | Out-Null; Release-Com $rwb }

    # Block 1's currency columns (C:M), read dynamically from header row 3
    $colMap = @{}   # currency code -> 1-based column index
    $c = 3   # C
    while ($reportWs.Cells.Item(3, $c).Text.Trim() -ne "") {
        $colMap[$reportWs.Cells.Item(3, $c).Text.Trim()] = $c
        $c++
    }

    # Find the row for $month in column B, or append after the last populated one
    $lastRow = Get-TrueLastDataRow $reportWs 2   # column B
    $targetRow = $null
    for ($r = 4; $r -le $lastRow; $r++) {
        if ($reportWs.Cells.Item($r, 2).Text.Trim() -eq $month) { $targetRow = $r; break }
    }
    if (-not $targetRow) { $targetRow = $lastRow + 1 }
    $reportWs.Cells.Item($targetRow, 2).Value2 = $month   # Block 1's own month column

    $filled = 0; $skipped = @()
    foreach ($code in $colMap.Keys) {
        if ($hkdRates.ContainsKey($code)) {
            $reportWs.Cells.Item($targetRow, $colMap[$code]).Value2 = $hkdRates[$code]
            $filled++
        } else {
            $skipped += $code
        }
    }

    # Block 2 (O:Z) formulas are always (re)written directly rather than copied from an adjacent row -
    # an earlier version copied from the row above, which broke the moment that row turned out to
    # already be frozen to a static value itself (copying a frozen row just copies its stale literal
    # text, not an auto-adjusting formula). Hardcoding the formula text is immune to that.
    $r = $targetRow
    $reportWs.Cells.Item($r, 15).Formula = "=B$r"                                              # O = month
    $reportWs.Cells.Item($r, 16).Formula = "=IF(C$r=`"`",`"`",ROUND(+C$r/K$r,4))"               # P = AUD
    $reportWs.Cells.Item($r, 17).Formula = "=IF(D$r=`"`",`"`",ROUND(+D$r/K$r,4))"               # Q = CAD
    $reportWs.Cells.Item($r, 18).Formula = "=IF(E$r=`"`",`"`",ROUND(+E$r/K$r,4))"               # R = CNY
    $reportWs.Cells.Item($r, 19).Formula = "=IF(F$r=`"`",`"`",ROUND(+F$r/K$r,4))"               # S = EUR
    $reportWs.Cells.Item($r, 20).Formula = "=IF(G$r=`"`",`"`",ROUND(+G$r/K$r,4))"               # T = GBP
    $reportWs.Cells.Item($r, 21).Formula = "=IF(H$r=`"`",`"`",ROUND(+H$r/K$r,4))"               # U = JPY
    $reportWs.Cells.Item($r, 22).Formula = "=IF(I$r=`"`",`"`",ROUND(+I$r/K$r,4))"               # V = NTD
    $reportWs.Cells.Item($r, 23).Formula = "=IF(J$r=`"`",`"`",ROUND(+J$r/K$r,4))"               # W = THB
    $reportWs.Cells.Item($r, 24).Formula = "=IF(K$r=`"`",`"`",ROUND(+K$r/K$r,4))"               # X = USD (always 1)
    $reportWs.Cells.Item($r, 25).Formula = "=IF(L$r=`"`",`"`",ROUND(+L$r/K$r,4))"               # Y = HKD
    $reportWs.Cells.Item($r, 26).Formula = "=IF(M$r=`"`",`"`",ROUND(+M$r/K$r,4))"               # Z = PEN

    Write-Host "[ExchangeRate] Filled $filled currency column(s) in Block 1 for $month at row $targetRow (Block 2 computes via its own formulas). $(if ($skipped) { "No source rate for: $($skipped -join ', ') - left as-is." })"
}

# Fills the Data sheet's shared formula columns (Q:S, U:AJ - everything except T, which is a plain DB
# marker, not a formula) for newly appended rows, ONE COLUMN AT A TIME: copy that column's row-2 master
# formula down over [startRow:endRow], force a recalc, then immediately freeze that column's result to
# static values before moving to the next column.
#
# Why one column at a time, not all at once: filling every formula column together and calculating once
# would recalculate the full dependency web across every new row simultaneously - this is exactly what
# caused Excel to hang for 10+ minutes earlier in this project. Freezing each column to values immediately
# after it calculates keeps every individual recalculation small, and later columns read already-resolved
# plain values instead of live formulas from earlier columns - functionally identical, far cheaper to
# compute.
#
# Order is dependency order, not left-to-right spreadsheet order, found by reading every formula in
# row 2 (e.g. Q references X, which sits to Q's right - naive left-to-right would fill Q before X
# exists). This function only ever gets called with the AA/AB-independent columns from Main (U,V,W,Q,
# R,S,Y,Z,AJ) - AA/AB are filled manually (see note in Main), and the AA/AB-dependent columns
# (AC/AD/AE/AF/AG/AH/AI) are filled by the separate Complete-FormulaColumns.ps1 script, run after the
# manual AA/AB step. Dependency notes for the columns THIS script's Main actually passes in:
#   U,V,W    <- raw columns only, no formula dependencies, can go first in any order
#              (X is deliberately excluded from this whole function - see note in Main)
#   Q        <- needs X (kept permanently live in the sheet, not filled by this function)
#   R, S     <- need Q
#   Y        <- needs U (and raw L,T)
#   Z, AJ    <- need W
# (AC/AF/AD/AE/AG/AH/AI's dependency notes live in Complete-FormulaColumns.ps1, where they're now filled)
function Set-FormulaColumnsFrozen($ws, [int]$startRow, [int]$endRow, [string[]]$columns) {
    if ($endRow -lt $startRow) { return }
    $xlPasteFormulas = -4123
    $xlPasteValues = -4163
    foreach ($col in $columns) {
        $target = "$col$startRow`:$col$endRow"
        $ws.Range("$col`2").Copy() | Out-Null
        $ws.Range($target).PasteSpecial($xlPasteFormulas) | Out-Null
        $ws.Calculate() | Out-Null
        # Freeze to static values via Copy+PasteSpecial(values) - NOT "$rng.Value2 = $rng.Value2"
        # (that self-reassignment trick throws "Specified cast is not valid" via COM interop; found
        # this the hard way). Copy+PasteSpecial is Excel's own native value-freeze mechanism.
        $ws.Range($target).Copy() | Out-Null
        $ws.Range($target).PasteSpecial($xlPasteValues) | Out-Null
        Write-Host "  [$col] filled, calculated, frozen"
    }
    $ws.Application.CutCopyMode = $false
}

# NOTE: AA/AB filling (formerly Set-MaxMinPriceColumns, a PowerShell-computed replacement for their
# MAXIFS/MINIFS formulas) and the "New Vendors" review-sheet scan (formerly Copy-NewVendorRowsToReviewSheet)
# both moved to Complete-FormulaColumns.ps1 (2026-08) - AA/AB are now filled manually by the report owner
# copying AA2:AB2's live formula down, and everything that depends on AA/AB (including the New Vendors
# scan's natural place in the column-fill order) runs afterward in that separate script. See its header
# for the full rationale.

# --- Main ---

$existingExcel = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($existingExcel) {
    Write-Host "Found $($existingExcel.Count) leftover EXCEL.EXE process(es) - closing them first to guarantee a clean COM session." -ForegroundColor Yellow
    Write-Host "(If you have unsaved work open in Excel, save it now - Ctrl+C to abort within 5 seconds.)" -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    $existingExcel | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

$excel = New-ExcelApp
$allNewRows = @()

try {
    # Safety check up front, before touching any source files: since -ReportTemplatePath now
    # auto-detects last month's own output (for continuity), an accidental re-run for a month
    # already present would silently duplicate rows. The month tag (col X) is always in the last
    # rows if present, since data is appended chronologically.
    if (-not $Force) {
        $checkWb = Open-Workbook $excel $ReportTemplatePath $true
        try {
            $checkData = $checkWb.Sheets.Item("MHK+STC+MZQ+MIL+RACO Data")
            $checkLastRow = Get-TrueLastDataRow $checkData 1
            $checkStart = [Math]::Max(2, $checkLastRow - 200)
            $recentX = Get-RangeArray $checkData $checkStart 24 $checkLastRow 24
            $alreadyPresent = $false
            for ($i = 0; $i -lt $recentX.GetLength(0); $i++) { if ($recentX[$i, 0] -eq $ReportMonth) { $alreadyPresent = $true; break } }
            if ($alreadyPresent) {
                throw "Data for $ReportMonth already appears to exist near the end of '$ReportTemplatePath' (found in column X). Re-running would likely duplicate rows. Pass -Force to proceed anyway."
            }
        } finally { $checkWb.Close($false) | Out-Null; Release-Com $checkWb }
    }

    $allNewRows += , (Process-MHK  $excel (Join-Path $WorkingFolder "MHK\MHK.xlsx"))
    $allNewRows += , (Process-STC  $excel (Join-Path $WorkingFolder "STC\STC.xlsx"))
    $allNewRows += , (Process-CumulativeDirect $excel (Join-Path $WorkingFolder "MZQ\MZQ.xlsx") "MHK" 2 $ReportMonth "MZQ")
    $allNewRows += , (Process-CumulativeDirect $excel (Join-Path $WorkingFolder "MIL\MIL.xlsx") "MIL" 2 $ReportMonth "MIL")
    $allNewRows += , (Process-RACO $excel (Join-Path $WorkingFolder "RACO\RACO.xlsx") $ReportMonth)
    $allNewRows += , (Process-TVS  $excel (Join-Path $WorkingFolder "TVS\TVS.xlsx") $ReportMonth)

    $mazalPath = Join-Path $WorkingFolder "MAZ\MAZ.xlsx"
    if (Test-Path $mazalPath) {
        $allNewRows += , (Process-MAZAL $excel $mazalPath $ReportMonth)
    } else {
        Write-Host "[MAZAL] $mazalPath not found - skipping (data not received yet this cycle)." -ForegroundColor Yellow
    }

    Write-Host "Copying report template to output path..."
    Copy-Item $ReportTemplatePath $OutputPath -Force
    $reportWb = Open-Workbook $excel $OutputPath $false
    try {
        $data = $reportWb.Sheets.Item("MHK+STC+MZQ+MIL+RACO Data")
        $usedLastRow = Get-UsedLastRow $data
        $trueLastRow = Get-TrueLastDataRow $data 1
        if ($trueLastRow -lt $usedLastRow) {
            Write-Host "Note: UsedRange reports row $usedLastRow but the true last data row (col A) is $trueLastRow - $($usedLastRow - $trueLastRow) trailing blank rows exist in the template. Appending after the true last data row." -ForegroundColor Yellow
        }

        $startRow = $trueLastRow + 1

        $r = $startRow
        foreach ($entityRows in $allNewRows) {
            $n = $entityRows.GetLength(0)
            if ($n -gt 0) {
                Set-RangeArray $data $r 1 $entityRows
                $r += $n
            }
        }
        $lastNewRow = $r - 1
        Write-Host "Appended rows $startRow to $lastNewRow into Report Data sheet."

        # Exchange Rate sheet must be updated BEFORE filling Q (EXCH_RATE looks up the target month's
        # row there via INDEX/MATCH) - order matters here.
        if ($ExchangeRatePath) {
            $erWs = $reportWb.Sheets.Item("Exchange Rate")
            Update-ExchangeRateSheet $excel $ExchangeRatePath $erWs $ReportMonth
        }

        # Fill Report's Q:S and the AA/AB-independent part of U:AJ formulas for the newly appended rows
        # only, one column at a time with an immediate freeze-to-values after each (see
        # Set-FormulaColumnsFrozen for why). Historical rows stay frozen from earlier months - this never
        # re-touches them. AA/AB and everything depending on them (AC/AD/AE/AF/AG/AH/AI, New Vendors scan)
        # are deliberately NOT done here - see the header note and the message printed below.
        if ($lastNewRow -ge $startRow) {
            Write-Host "Filling Q:S/U:AJ (AA/AB-independent columns) for rows $startRow to $lastNewRow (one column at a time)..."
            Set-FormulaColumnsFrozen $data $startRow $lastNewRow @("U", "V", "W", "Q", "R", "S", "Y", "Z", "AJ")
        }

        $reportWb.Save()
        Write-Host "Saved: $OutputPath" -ForegroundColor Green
    } finally { $reportWb.Close($false) | Out-Null; Release-Com $reportWb }
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    exit 1
}
finally {
    if ($excel) { $excel.Quit(); Release-Com $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "=== Done with the automated part. Manual step needed next: ===" -ForegroundColor Cyan
Write-Host "  1. Open the report in Excel, copy AA2:AB2 down to rows $startRow-$lastNewRow, and calculate." -ForegroundColor Cyan
Write-Host "  2. Once confirmed correct, run: .\Complete-FormulaColumns.ps1 -ReportMonth `"$ReportMonth`"" -ForegroundColor Cyan
Write-Host "     (fills AC/AD/AE/AF/AG/AH/AI, freezes them, and builds the New Vendors review sheet)" -ForegroundColor Cyan
