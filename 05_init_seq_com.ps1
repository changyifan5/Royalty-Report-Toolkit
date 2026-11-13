param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [int]$HeaderRow = 6,
    [int]$DataStartRow = 7,
    [string]$SeqHeader = "__原序",
    [int]$SheetIndex = 0,
    [int]$ChunkSize = 50000
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$xlUp = -4162
$xlToLeft = -4159
$xlFillDefault = 0

function Release-ComObjectSafe($obj) {
    if ($null -ne $obj) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj)
    }
}

function Get-ColumnLetter([int]$col) {
    $letters = ""
    while ($col -gt 0) {
        $mod = ($col - 1) % 26
        $letters = [char](65 + $mod) + $letters
        $col = [int](($col - $mod) / 26)
    }
    return $letters
}

function Find-SeqColumn($Worksheet, [int]$HeaderRow, [string]$SeqHeader, [int]$LastHeaderCol) {
    for ($c = 1; $c -le $LastHeaderCol; $c++) {
        $text = [string]$Worksheet.Cells.Item($HeaderRow, $c).Text
        if ($text.Trim() -eq $SeqHeader) {
            return $c
        }
    }
    return 0
}

function Test-SeqColumnEmpty($Worksheet, [int]$SeqCol, [int]$DataStartRow, [int]$LastRow) {
    if ($LastRow -lt $DataStartRow) { return $true }
    $first = $Worksheet.Cells.Item($DataStartRow, $SeqCol).Value2
    $last = $Worksheet.Cells.Item($LastRow, $SeqCol).Value2
    return ($null -eq $first -or [string]$first -eq "") -and ($null -eq $last -or [string]$last -eq "")
}

function Write-SeqColumnValues($Worksheet, [int]$DataStartRow, [int]$LastRow, [int]$SeqCol, [int]$ChunkSize) {
    $rowCount = $LastRow - $DataStartRow + 1
    if ($rowCount -le 0) { return }

    if ($rowCount -eq 1) {
        $Worksheet.Cells.Item($DataStartRow, $SeqCol).Value2 = 1
        return
    }

    if ($rowCount -le $ChunkSize) {
        try {
            $seedTop = $Worksheet.Cells.Item($DataStartRow, $SeqCol)
            $seedBottom = $Worksheet.Cells.Item($DataStartRow + 1, $SeqCol)
            $seedTop.Value2 = 1
            $seedBottom.Value2 = 2
            $seedRange = $Worksheet.Range($seedTop, $seedBottom)
            $destRange = $Worksheet.Range(
                $Worksheet.Cells.Item($DataStartRow, $SeqCol),
                $Worksheet.Cells.Item($LastRow, $SeqCol)
            )
            $seedRange.AutoFill($destRange, $xlFillDefault) | Out-Null
            return
        }
        catch {
            Write-Output "WARN,autofill_failed,$($_.Exception.Message)"
        }
    }

    for ($start = 0; $start -lt $rowCount; $start += $ChunkSize) {
        $chunkLen = [Math]::Min($ChunkSize, $rowCount - $start)
        $rowBegin = $DataStartRow + $start
        $rowEnd = $rowBegin + $chunkLen - 1
        $values = New-Object 'object[,]' $chunkLen, 1
        for ($i = 0; $i -lt $chunkLen; $i++) {
            $values[$i, 0] = [double]($start + $i + 1)
        }
        $range = $Worksheet.Range(
            $Worksheet.Cells.Item($rowBegin, $SeqCol),
            $Worksheet.Cells.Item($rowEnd, $SeqCol)
        )
        $range.Value2 = $values
    }
}

$Excel = $null
$wb = $null
$saved = $false

try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Excel.ScreenUpdating = $false
    $Excel.EnableEvents = $false

    $wb = $Excel.Workbooks.Open($Path, $null, $false)
    if ($wb.Worksheets.Count -le $SheetIndex) {
        Write-Output "ERROR,no_sheet"
        exit 1
    }

    $ws = $wb.Worksheets.Item($SheetIndex + 1)

    $lastHeaderCol = $ws.Cells.Item($HeaderRow, $ws.Columns.Count).End($xlToLeft).Column
    if ($lastHeaderCol -lt 1) { $lastHeaderCol = 1 }

    $lastRow = $ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row
    if ($lastRow -lt $DataStartRow) {
        Write-Output "ERROR,no_data"
        exit 3
    }

    $existingSeqCol = Find-SeqColumn $ws $HeaderRow $SeqHeader ([Math]::Max($lastHeaderCol, 1))
    # 也扫描表头右侧一列（上次失败可能只写入了 __原序 表头）
    if ($existingSeqCol -eq 0) {
        $maybeCol = $lastHeaderCol + 1
        $text = [string]$ws.Cells.Item($HeaderRow, $maybeCol).Text
        if ($text.Trim() -eq $SeqHeader) {
            $existingSeqCol = $maybeCol
        }
    }
    $refilled = $false

    if ($existingSeqCol -gt 0) {
        if (-not (Test-SeqColumnEmpty $ws $existingSeqCol $DataStartRow $lastRow)) {
            $colLetter = Get-ColumnLetter $existingSeqCol
            Write-Output "ERROR,already_has_seq,$colLetter"
            exit 2
        }
        $seqCol = $existingSeqCol
        $refilled = $true
        $colLetter = Get-ColumnLetter $seqCol
        Write-Output "RECOVER,refill_empty_seq,$colLetter"
    }
    else {
        $seqCol = $lastHeaderCol + 1
        $ws.Cells.Item($HeaderRow, $seqCol).Value2 = $SeqHeader
    }

    $rowCount = $lastRow - $DataStartRow + 1
    Write-SeqColumnValues $ws $DataStartRow $lastRow $seqCol $ChunkSize

    $wb.Save() | Out-Null
    $saved = $true
    $colLetter = Get-ColumnLetter $seqCol
    if ($refilled) {
        Write-Output "STATS,$($ws.Name),$colLetter,$rowCount,refilled"
    }
    else {
        Write-Output "STATS,$($ws.Name),$colLetter,$rowCount"
    }
}
catch {
    Write-Output "ERROR,$($_.Exception.Message)"
    exit 1
}
finally {
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
