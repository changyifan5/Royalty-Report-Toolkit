# COM：B 列 AutoFilter 扫描（09 / 11 共用）
param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$FilterMode,
    [int]$SheetIndex = 0,
    [int]$HeaderRow = 6,
    [int]$DataStartRow = 7,
    [int]$DataEndRow = 0,
    [int]$FilterCol = 2
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
. "$PSScriptRoot\com_progress_common.ps1"
. "$PSScriptRoot\excel_filter_common.ps1"

function Release-ComObjectSafe($obj) {
    if ($null -ne $obj) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj)
    }
}

if (-not (Test-Path -LiteralPath $Path)) {
    Write-ProgressLine "ERROR,file_not_found,$Path"
    exit 1
}

$Excel = $null
$wb = $null

try {
    Write-ProgressLine "PROGRESS,stage,scan,0"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.ScreenUpdating = $false
    $Excel.EnableEvents = $false

    $wb = $Excel.Workbooks.Open($Path, 0, $true, 5, "", "", $true, "", "", $false, $false, $false)
    if ($wb.Worksheets.Count -le $SheetIndex) {
        Write-ProgressLine "ERROR,no_sheet"
        exit 1
    }

    $ws = $wb.Worksheets.Item($SheetIndex + 1)
    $probeLastRow = Get-LastDataRow $ws $DataStartRow
    if ($DataEndRow -gt 0) {
        $probeLastRow = [Math]::Min($probeLastRow, $DataEndRow)
    }
    if ($probeLastRow -lt $DataStartRow) {
        Write-ProgressLine "ROWS,"
        Write-ProgressLine "STATS,0,0"
        exit 0
    }

    $total = [int]($probeLastRow - $DataStartRow + 1)
    Write-ProgressLine "PROGRESS,scan,0,$total"
    Write-ProgressLine "INFO,filter_mode,$FilterMode"
    Write-ProgressLine "INFO,filter_col,$FilterCol"

    $scanResult = Invoke-ScanByColumnFilter $ws $HeaderRow $DataStartRow $probeLastRow $FilterCol $FilterMode {
        param([int]$Done, [int]$Total)
        Write-ProgressLine "PROGRESS,scan,$Done,$Total"
    }
    $rows = @($scanResult[0])
    $filterColor = [int]$scanResult[1]
    $lastCol = [int]$scanResult[2]
    $scanMethod = [string]$scanResult[3]

    $sw.Stop()
    Write-ProgressLine "PROGRESS,scan,$total,$total"
    Write-ProgressLine "TIMING,scan_ms,$($sw.ElapsedMilliseconds)"
    Write-ProgressLine "INFO,data_rows,$total"
    Write-ProgressLine "INFO,last_col,$lastCol"
    Write-ProgressLine "INFO,scan_method,$scanMethod"
    if ($filterColor -gt 0) {
        Write-ProgressLine "INFO,filter_color,$filterColor"
    }
    if ($rows.Count -gt 0) {
        Write-ProgressLine "ROWS,$($rows -join ',')"
    }
    else {
        Write-ProgressLine "ROWS,"
    }
    Write-ProgressLine "STATS,$($rows.Count),$total"
}
catch {
    Write-ProgressLine "ERROR,$($_.Exception.Message)"
    exit 1
}
finally {
    if ($null -ne $wb) {
        $wb.Close($false) | Out-Null
        Release-ComObjectSafe $wb
    }
    if ($null -ne $Excel) {
        $Excel.Quit() | Out-Null
        Release-ComObjectSafe $Excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
