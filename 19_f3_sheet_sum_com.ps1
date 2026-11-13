# 19 F3页签汇总：COM 读取各页签 F3（含公式计算值，按需计算）
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
. "$PSScriptRoot\com_progress_common.ps1"

function Convert-ComNumber($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return $null }
    if ($Value -is [string]) {
        $text = $Value.Trim()
        if ($text -eq "") { return $null }
        if ($text.StartsWith("#")) { return $null }
        $parsed = 0.0
        if ([double]::TryParse($text.Replace(",", ""), [ref]$parsed)) {
            return $parsed
        }
        return $null
    }
    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        return [double]$Value
    }
    return $null
}

function Get-RawDisplay($Value) {
    if ($null -eq $Value) { return "(空)" }
    return [string]$Value
}

function Read-F3Entry($Worksheet, [int]$Row, [int]$Col) {
    $name = [string]$Worksheet.Name
    $cell = $Worksheet.Cells.Item($Row, $Col)
    $hasFormula = $false
    try {
        $hasFormula = [bool]$cell.HasFormula
    }
    catch {
        $hasFormula = $false
    }

    $raw = $cell.Value2
    $num = Convert-ComNumber $raw
    return [ordered]@{
        name        = $name
        value       = $num
        has_formula = $hasFormula
        invalid     = ($null -eq $num)
        raw         = (Get-RawDisplay $raw)
        cell        = $cell
        needs_calc  = ($hasFormula -and $null -eq $num)
    }
}

$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$paths = @($manifest.paths)
$skipName = [string]$manifest.skip_sheet_name
$row = [int]$manifest.cell_row
$col = [int]$manifest.cell_col
$resultPath = [string]$manifest.result_path

$excel = $null
$filesOut = @()
$fileTotal = $paths.Count
$fileDone = 0

try {
    Write-ProgressLine "PROGRESS,stage,open,0"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.Calculation = -4135

    foreach ($path in $paths) {
        $fileDone++
        Write-ProgressLine "PROGRESS,files,$fileDone,$fileTotal"
        $resolved = [System.IO.Path]::GetFullPath([string]$path)
        if (-not (Test-Path -LiteralPath $resolved)) {
            Write-ProgressLine "ERROR,找不到文件: $resolved"
            continue
        }

        $wb = $null
        $fileEntry = [ordered]@{
            path    = $resolved
            sheets  = @()
            skipped = @()
        }
        try {
            $wb = $excel.Workbooks.Open(
                $resolved,
                0,
                $true,
                $null,
                $null,
                $null,
                $true,
                $null,
                $null,
                $false,
                $false,
                $null,
                $false,
                $false,
                $false
            )

            $pending = New-Object System.Collections.Generic.List[object]
            foreach ($ws in @($wb.Worksheets)) {
                $name = [string]$ws.Name
                if ($name -eq $skipName) {
                    $fileEntry.skipped += $name
                    continue
                }

                $entry = Read-F3Entry $ws $row $col
                if ($entry.needs_calc) {
                    $pending.Add($entry) | Out-Null
                }
                else {
                    $fileEntry.sheets += [ordered]@{
                        name        = $entry.name
                        value       = $entry.value
                        has_formula = $entry.has_formula
                        invalid     = $entry.invalid
                        raw         = $entry.raw
                    }
                }
            }

            if ($pending.Count -gt 0) {
                try { $null = $wb.Calculate() } catch { }
                foreach ($entry in $pending) {
                    $raw = $entry.cell.Value2
                    $num = Convert-ComNumber $raw
                    $fileEntry.sheets += [ordered]@{
                        name        = $entry.name
                        value       = $num
                        has_formula = $entry.has_formula
                        invalid     = ($null -eq $num)
                        raw         = (Get-RawDisplay $raw)
                    }
                }
            }
        }
        finally {
            if ($wb) {
                $wb.Close($false) | Out-Null
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            }
        }
        $filesOut += ,$fileEntry
    }

    $sw.Stop()
    Write-ProgressLine "TIMING,read_ms,$($sw.ElapsedMilliseconds)"
    $json = ($filesOut | ConvertTo-Json -Depth 8 -Compress)
    if ($resultPath) {
        [System.IO.File]::WriteAllText($resultPath, $json, [System.Text.UTF8Encoding]::new($false))
        Write-ProgressLine "RESULT_FILE,$resultPath"
    }
    else {
        Write-ProgressLine "RESULT_JSON,$json"
    }
    Write-ProgressLine "STATS,files,$fileTotal"
}
catch {
    Write-ProgressLine "ERROR,$($_.Exception.Message)"
    exit 1
}
finally {
    if ($excel) {
        $excel.Quit() | Out-Null
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
