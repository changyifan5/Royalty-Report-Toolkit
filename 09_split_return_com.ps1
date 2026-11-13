# 09 split-return (COM): read manifest, batch-copy rows (optimized)
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$xlToLeft = -4159
$xlUp = -4162
$xlPasteColumnWidths = 7
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105

function Get-ExcelFileFormat([string]$Path) {
    switch -Regex ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '\.xlsm$' { return 52 }
        '\.xls$'  { return 56 }
        '\.xltx$' { return 54 }
        '\.xltm$' { return 55 }
        default   { return 51 }
    }
}

function Clear-FileReadOnly([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $item = Get-Item -LiteralPath $Path
    if ($item.IsReadOnly) {
        $item.IsReadOnly = $false
    }
}

function Save-WorkbookForce([object]$Workbook, [string]$Path) {
    $format = Get-ExcelFileFormat $Path
    Clear-FileReadOnly $Path
    if ($Workbook.ReadOnly) {
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force
        }
        $Workbook.SaveAs($Path, $format) | Out-Null
    }
    else {
        $Workbook.Save() | Out-Null
    }
    if ($Workbook.ReadOnly) {
        throw "Workbook is still read-only after save: $Path"
    }
}

function Release-ComObjectSafe($obj) {
    if ($null -ne $obj) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj)
    }
}

function Clear-CopyMode($Excel) {
    if ($null -ne $Excel) {
        $Excel.CutCopyMode = $false
    }
}

function Get-SourceLastCol($SourceSheet, [int]$HeaderRow, [int]$FallbackCol) {
    if ($FallbackCol -gt 0) { return $FallbackCol }
    $lastCol = $SourceSheet.Cells.Item($HeaderRow, $SourceSheet.Columns.Count).End($xlToLeft).Column
    if ($lastCol -lt 1) { return 1 }
    return $lastCol
}

function Copy-HeaderLayout($Excel, $SourceSheet, $DestSheet, [int]$HeaderRows, [int]$LastCol) {
    if ($LastCol -lt 1) { $LastCol = 1 }

    $srcHeader = $SourceSheet.Range(
        $SourceSheet.Cells.Item(1, 1),
        $SourceSheet.Cells.Item($HeaderRows, $LastCol)
    )
    $dstHeader = $DestSheet.Range(
        $DestSheet.Cells.Item(1, 1),
        $DestSheet.Cells.Item($HeaderRows, $LastCol)
    )
    $srcHeader.Copy($dstHeader) | Out-Null

    $srcCols = $SourceSheet.Range(
        $SourceSheet.Cells.Item(1, 1),
        $SourceSheet.Cells.Item(1, $LastCol)
    ).EntireColumn
    $dstCols = $DestSheet.Range(
        $DestSheet.Cells.Item(1, 1),
        $DestSheet.Cells.Item(1, $LastCol)
    ).EntireColumn
    $srcCols.Copy() | Out-Null
    $dstCols.PasteSpecial($xlPasteColumnWidths) | Out-Null

    for ($r = 1; $r -le $HeaderRows; $r++) {
        $DestSheet.Rows.Item($r).RowHeight = $SourceSheet.Rows.Item($r).RowHeight
    }
    Clear-CopyMode $Excel
}

function Get-RowGroups([int[]]$Rows) {
    if ($Rows.Count -eq 0) { return @() }
    $groups = New-Object System.Collections.Generic.List[object]
    $start = $Rows[0]
    $prev = $Rows[0]
    for ($i = 1; $i -lt $Rows.Count; $i++) {
        $r = $Rows[$i]
        if ($r -ne ($prev + 1)) {
            $groups.Add(@($start, $prev))
            $start = $r
        }
        $prev = $r
    }
    $groups.Add(@($start, $prev))
    return $groups
}

function Convert-ManifestRowGroups($ManifestGroups) {
    $groups = New-Object System.Collections.Generic.List[object]
    foreach ($g in @($ManifestGroups)) {
        if ($g -is [System.Array] -or $g -is [Object[]]) {
            $groups.Add(@([int]$g[0], [int]$g[1]))
        }
        else {
            $groups.Add(@([int]$g.'0', [int]$g.'1'))
        }
    }
    return $groups
}

function Copy-RowGroups($Excel, $SourceSheet, $DestSheet, $RowGroups, [int]$LastCol, [int]$DataStartRow) {
    if ($LastCol -lt 1) { $LastCol = 1 }
    $destRow = $DataStartRow
    $total = @($RowGroups).Count
    $idx = 0

    foreach ($g in @($RowGroups)) {
        $idx++
        $startRow = [int]$g[0]
        $endRow = [int]$g[1]
        $rowCount = $endRow - $startRow + 1
        $destEnd = $destRow + $rowCount - 1

        $srcRange = $SourceSheet.Range(
            $SourceSheet.Cells.Item($startRow, 1),
            $SourceSheet.Cells.Item($endRow, $LastCol)
        )
        $dstRange = $DestSheet.Range(
            $DestSheet.Cells.Item($destRow, 1),
            $DestSheet.Cells.Item($destEnd, $LastCol)
        )
        $srcRange.Copy($dstRange) | Out-Null
        $destRow += $rowCount

        if ($total -ge 10 -and ($idx % 10 -eq 0 -or $idx -eq $total)) {
            Write-Output "PROGRESS,copy_groups,$idx,$total"
        }
    }
    Clear-CopyMode $Excel
}

if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-Output "ERROR,manifest_not_found,$ManifestPath"
    exit 1
}

$manifest = Get-Content -LiteralPath $ManifestPath -Encoding UTF8 -Raw | ConvertFrom-Json
$path = [string]$manifest.path
$destSheetName = [System.Text.Encoding]::UTF8.GetString(
    [Convert]::FromBase64String([string]$manifest.dest_sheet_name_b64)
)
$sourceSheetIndex = [int]$manifest.source_sheet_index
$headerRows = [int]$manifest.header_rows
$dataStartRow = [int]$manifest.data_start_row
$lastCol = 0
if ($manifest.PSObject.Properties.Name -contains 'last_col') {
    $lastCol = [int]$manifest.last_col
}

$dataRows = @()
if ($manifest.PSObject.Properties.Name -contains 'data_rows') {
    $dataRows = @($manifest.data_rows | ForEach-Object { [int]$_ })
}
else {
    $dataRows = @($manifest.red_rows | ForEach-Object { [int]$_ })
}

$rowGroups = @()
if ($manifest.PSObject.Properties.Name -contains 'row_groups') {
    $rowGroups = Convert-ManifestRowGroups $manifest.row_groups
}
else {
    $rowGroups = Get-RowGroups $dataRows
}

$appendToExisting = $false
if ($manifest.PSObject.Properties.Name -contains 'append_to_existing') {
    $appendToExisting = [bool]$manifest.append_to_existing
}

if (-not (Test-Path -LiteralPath $path)) {
    Write-Output "ERROR,file_not_found,$path"
    exit 1
}

Clear-FileReadOnly $path

$Excel = $null
$wb = $null
$saved = $false
$oldCalc = $null

try {
    $swOpen = [System.Diagnostics.Stopwatch]::StartNew()
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.ScreenUpdating = $false
    $Excel.EnableEvents = $false
    $Excel.Interactive = $false
    $oldCalc = $Excel.Calculation
    $Excel.Calculation = $xlCalculationManual

    $wb = $Excel.Workbooks.Open($path, 0, $false, 5, "", "", $true, "", "", $false, $false, $false)
    $swOpen.Stop()
    Write-Output "TIMING,open_ms,$($swOpen.ElapsedMilliseconds)"

    if ($wb.ReadOnly) {
        Write-Output "WARN,opened_readonly,trying_saveas"
    }
    if ($wb.Worksheets.Count -le $sourceSheetIndex) {
        Write-Output "ERROR,no_sheet"
        exit 1
    }

    $src = $wb.Worksheets.Item($sourceSheetIndex + 1)
    $sourceName = $src.Name
    $lastCol = Get-SourceLastCol $src $headerRows $lastCol
    Write-Output "INFO,last_col,$lastCol"
    Write-Output "INFO,row_groups,$(@($rowGroups).Count)"
    Write-Output "INFO,data_rows,$($dataRows.Count)"

    $dest = $null
    if ($appendToExisting) {
        foreach ($ws in @($wb.Worksheets)) {
            if ($ws.Name -eq $destSheetName) {
                $dest = $ws
                break
            }
        }
        if ($null -eq $dest) {
            Write-Output "ERROR,append_sheet_missing,$destSheetName"
            exit 1
        }
    }
    else {
        foreach ($ws in @($wb.Worksheets)) {
            if ($ws.Name -eq $destSheetName) {
                $ws.Delete()
                break
            }
        }

        $dest = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
        $dest.Name = $destSheetName

        $swHeader = [System.Diagnostics.Stopwatch]::StartNew()
        Copy-HeaderLayout $Excel $src $dest $headerRows $lastCol
        $swHeader.Stop()
        Write-Output "TIMING,header_ms,$($swHeader.ElapsedMilliseconds)"
    }

    $destStartRow = $dataStartRow
    if ($appendToExisting) {
        $lastRow = $dest.Cells.Item($dest.Rows.Count, 1).End($xlUp).Row
        if ($lastRow -ge $dataStartRow) {
            $destStartRow = $lastRow + 1
        }
    }

    $swCopy = [System.Diagnostics.Stopwatch]::StartNew()
    Copy-RowGroups $Excel $src $dest $rowGroups $lastCol $destStartRow
    $swCopy.Stop()
    Write-Output "TIMING,copy_ms,$($swCopy.ElapsedMilliseconds)"

    $swSave = [System.Diagnostics.Stopwatch]::StartNew()
    Save-WorkbookForce $wb $path
    $swSave.Stop()
    Write-Output "TIMING,save_ms,$($swSave.ElapsedMilliseconds)"

    $saved = $true
    Write-Output "SAVED,$path"
    Write-Output "STATS,$destSheetName,$sourceName,$($dataRows.Count)"
}
catch {
    Write-Output "ERROR,$($_.Exception.Message)"
    exit 1
}
finally {
    Clear-CopyMode $Excel
    if ($null -ne $Excel -and $null -ne $oldCalc) {
        try { $Excel.Calculation = $oldCalc } catch { }
    }
    if ($null -ne $wb) {
        $wb.Close($saved) | Out-Null
        Release-ComObjectSafe $wb
    }
    if ($null -ne $Excel) {
        $Excel.Quit() | Out-Null
        Release-ComObjectSafe $Excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
