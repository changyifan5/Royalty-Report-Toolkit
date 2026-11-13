# 18 新建公式列：在指定列后插入一列并写入 Excel 公式（可多页签）
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
. "$PSScriptRoot\com_progress_common.ps1"

$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105
$xlToRight = -4161
$xlUp = -4162

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
}

function Release-ComObjectSafe($obj) {
    if ($null -ne $obj) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj)
    }
}

function Get-WorksheetByName($Workbook, [string]$Name) {
    for ($i = 1; $i -le $Workbook.Worksheets.Count; $i++) {
        $candidate = $Workbook.Worksheets.Item($i)
        if ($candidate.Name -eq $Name) {
            return $candidate
        }
        Release-ComObjectSafe $candidate
    }
    return $null
}

if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-Output "ERROR,manifest_not_found,$ManifestPath"
    exit 1
}

$manifest = Get-Content -LiteralPath $ManifestPath -Encoding UTF8 -Raw | ConvertFrom-Json
$path = [string]$manifest.path
$headerRow = [int]$manifest.header_row
$xlSolid = 1
$jobs = @($manifest.jobs)

if (-not (Test-Path -LiteralPath $path)) {
    Write-Output "ERROR,file_not_found,$path"
    exit 1
}
if ($jobs.Count -lt 1) {
    Write-Output "ERROR,no_jobs"
    exit 1
}

Clear-FileReadOnly $path

$Excel = $null
$wb = $null
$saved = $false
$oldCalc = $null
$totalWritten = 0
$doneSheets = 0

try {
    Write-ProgressLine "PROGRESS,stage,open,0"
    $swOpen = [System.Diagnostics.Stopwatch]::StartNew()
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.ScreenUpdating = $false
    $Excel.EnableEvents = $false
    $oldCalc = $Excel.Calculation
    $Excel.Calculation = $xlCalculationManual
    $wb = $Excel.Workbooks.Open($path, 0, $false, 5, "", "", $true, "", "", $false, $false, $false)
    $swOpen.Stop()
    Write-ProgressLine "TIMING,open_ms,$($swOpen.ElapsedMilliseconds)"

    $jobCount = $jobs.Count
    $step = Get-ProgressStep $jobCount
    Write-ProgressLine "PROGRESS,stage,insert,0"
    $swFormula = [System.Diagnostics.Stopwatch]::StartNew()

    for ($j = 0; $j -lt $jobCount; $j++) {
        $job = $jobs[$j]
        $sheetName = [string]$job.sheet_name
        $afterCol = [int]$job.after_col
        $newHeader = [string]$job.new_header
        $formulaText = ""
        if ($job.PSObject.Properties.Name -contains 'formula' -and $null -ne $job.formula) {
            $formulaText = [string]$job.formula
        }
        $dataStart = [int]$job.data_start_row
        $dataEnd = [int]$job.data_end_row
        $formulaFirst = [string]$job.formula_first
        $fillBgr = $null
        if ($job.PSObject.Properties.Name -contains 'fill_bgr' -and $null -ne $job.fill_bgr -and [string]$job.fill_bgr -ne "") {
            $fillBgr = [int]$job.fill_bgr
        }

        $ws = Get-WorksheetByName $wb $sheetName
        if ($null -eq $ws) {
            Write-ProgressLine "SKIP,$sheetName"
            continue
        }

        $insertAt = $afterCol + 1
        if ($insertAt -lt 1) { $insertAt = 1 }
        $insertRange = $ws.Columns.Item($insertAt)
        $insertRange.Insert($xlToRight) | Out-Null
        Release-ComObjectSafe $insertRange
        $ws.Cells.Item($headerRow, $insertAt).Value2 = $newHeader
        if ($afterCol -ge 1) {
            try {
                $ws.Columns.Item($insertAt).ColumnWidth = $ws.Columns.Item($afterCol).ColumnWidth
            }
            catch { }
        }

        $written = 0
        if ($dataEnd -lt $dataStart) {
            $dataEnd = $dataStart - 1
        }
        try {
            $probeLast = [int]$ws.Cells.Item($ws.Rows.Count, [Math]::Max(1, $afterCol)).End($xlUp).Row
            if ($probeLast -gt $dataEnd) { $dataEnd = $probeLast }
        }
        catch { }
        try {
            $probeLastA = [int]$ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row
            if ($probeLastA -gt $dataEnd) { $dataEnd = $probeLastA }
        }
        catch { }

        if (-not [string]::IsNullOrWhiteSpace($formulaText) -and $dataEnd -ge $dataStart) {
            $totalFormula = $dataEnd - $dataStart + 1
            Write-ProgressLine "PROGRESS,stage,formula,0"
            Write-ProgressLine "PROGRESS,formula,0,$totalFormula"

            $firstCell = $ws.Cells.Item($dataStart, $insertAt)
            $firstCell.Formula = $formulaFirst
            $formulaR1C1 = [string]$firstCell.FormulaR1C1
            if ([string]::IsNullOrWhiteSpace($formulaR1C1)) {
                throw "formula_r1c1_empty"
            }
            $chunkStep = Get-ProgressStep $totalFormula
            $chunkSize = [Math]::Max(5000, $chunkStep)
            for ($r = $dataStart; $r -le $dataEnd; $r += $chunkSize) {
                $chunkEnd = [Math]::Min($dataEnd, $r + $chunkSize - 1)
                $chunk = $ws.Range(
                    $ws.Cells.Item($r, $insertAt),
                    $ws.Cells.Item($chunkEnd, $insertAt)
                )
                $chunk.FormulaR1C1 = $formulaR1C1
                Release-ComObjectSafe $chunk
                $doneRows = $chunkEnd - $dataStart + 1
                if ($r -eq $dataStart -or $chunkEnd -eq $dataEnd -or ($doneRows % $chunkStep -eq 0)) {
                    Write-ProgressLine "PROGRESS,formula,$doneRows,$totalFormula"
                }
            }
            Write-ProgressLine "PROGRESS,formula,$totalFormula,$totalFormula"
            $written = $totalFormula
            Release-ComObjectSafe $firstCell
        }

        if ($null -ne $fillBgr) {
            $headerCell = $ws.Cells.Item($headerRow, $insertAt)
            $headerCell.Interior.Pattern = $xlSolid
            $headerCell.Interior.Color = $fillBgr
            Release-ComObjectSafe $headerCell
        }

        $totalWritten += $written
        $doneSheets += 1
        Release-ComObjectSafe $ws

        $done = $j + 1
        if ($done -eq 1 -or $done -eq $jobCount -or ($done % $step -eq 0)) {
            Write-ProgressLine "PROGRESS,sheets,$done,$jobCount"
        }
    }

    $swFormula.Stop()
    Write-ProgressLine "TIMING,formula_ms,$($swFormula.ElapsedMilliseconds)"

    Write-ProgressLine "PROGRESS,stage,save,0"
    $swSave = [System.Diagnostics.Stopwatch]::StartNew()
    if ($null -ne $oldCalc) {
        $Excel.Calculation = $oldCalc
    }
    else {
        $Excel.Calculation = $xlCalculationAutomatic
    }
    Save-WorkbookForce $wb $path
    $saved = $true
    $swSave.Stop()
    Write-ProgressLine "TIMING,save_ms,$($swSave.ElapsedMilliseconds)"

    Write-ProgressLine "STATS,ok,$totalWritten,$doneSheets"
}
catch {
    Write-Output "ERROR,$($_.Exception.Message)"
    exit 1
}
finally {
    if ($null -ne $Excel -and $null -ne $oldCalc) {
        try { $Excel.Calculation = $oldCalc } catch { }
        try { $Excel.ScreenUpdating = $true } catch { }
        try { $Excel.EnableEvents = $true } catch { }
        try { $Excel.DisplayAlerts = $true } catch { }
    }
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch { }
        Release-ComObjectSafe $wb
    }
    if ($null -ne $Excel) {
        try { $Excel.Quit() } catch { }
        Release-ComObjectSafe $Excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not $saved) {
    Write-Output "ERROR,not_saved"
    exit 1
}
