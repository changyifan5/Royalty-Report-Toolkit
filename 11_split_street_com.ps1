# 11 街版拆分 (COM): 公式列 + 新建「街版」页签
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$xlToLeft = -4159
$xlPasteColumnWidths = 7
$xlCalculationManual = -4135

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

function Clear-CopyMode($Excel) {
    if ($null -ne $Excel) {
        $Excel.CutCopyMode = $false
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

function Get-HeaderLastCol($Worksheet, [int]$HeaderRow, [int]$Limit = 500) {
    $last = 0
    for ($c = 1; $c -le $Limit; $c++) {
        $text = [string]$Worksheet.Cells.Item($HeaderRow, $c).Text
        if ($text.Trim()) {
            $last = $c
        }
    }
    return $last
}

function Find-HeaderCol($Worksheet, [int]$HeaderRow, [string]$Name, [int]$Limit = 500) {
    for ($c = 1; $c -le $Limit; $c++) {
        $text = [string]$Worksheet.Cells.Item($HeaderRow, $c).Text
        if ($text.Trim() -eq $Name) {
            return $c
        }
    }
    return 0
}

function Get-NumericValue($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        return [double]$Value
    }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $text = $text.Trim().Replace(",", "")
    $parsed = 0.0
    if ([double]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }
    return $null
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

function Set-SplitFormulas(
    $Worksheet,
    [int]$Row,
    [int]$SourceCol,
    [int]$PaiCol,
    [int]$JieCol,
    [int]$OutCol1,
    [int]$OutCol2
) {
    $gL = Get-ColumnLetter $SourceCol
    $paiL = Get-ColumnLetter $PaiCol
    $jieL = Get-ColumnLetter $JieCol
    $c1L = Get-ColumnLetter $OutCol1
    $denom = "($paiL$Row+$jieL$Row)"
    $f1 = "=IF($denom=0,"""",$gL$Row*$paiL$Row/$denom)"
    $f2 = "=IF($denom=0,"""",$gL$Row-$c1L$Row)"
    $Worksheet.Cells.Item($Row, $OutCol1).Formula = $f1
    $Worksheet.Cells.Item($Row, $OutCol2).Formula = $f2
    $srcFmt = $Worksheet.Cells.Item($Row, $SourceCol).NumberFormat
    $Worksheet.Cells.Item($Row, $OutCol1).NumberFormat = $srcFmt
    $Worksheet.Cells.Item($Row, $OutCol2).NumberFormat = $srcFmt
}

function Get-StreetGValue($Worksheet, [int]$Row, [int]$SourceCol, [int]$PaiCol, [int]$JieCol) {
    $g = Get-NumericValue $Worksheet.Cells.Item($Row, $SourceCol).Value2
    if ($null -eq $g) { return $null }
    $pai = Get-NumericValue $Worksheet.Cells.Item($Row, $PaiCol).Value2
    $jie = Get-NumericValue $Worksheet.Cells.Item($Row, $JieCol).Value2
    if ($null -eq $pai) { $pai = 0.0 }
    if ($null -eq $jie) { $jie = 0.0 }
    $denom = $pai + $jie
    if ($denom -eq 0) { return $null }
    return [double]($g * $jie / $denom)
}

if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-Output "ERROR,manifest_not_found,$ManifestPath"
    exit 1
}

$manifest = Get-Content -LiteralPath $ManifestPath -Encoding UTF8 -Raw | ConvertFrom-Json
$path = [string]$manifest.path
$sheetIndex = [int]$manifest.sheet_index
$headerRow = [int]$manifest.header_row
$dataStartRow = [int]$manifest.data_start_row
$sourceCol = [int]$manifest.source_col
$streetSheet = [string]$manifest.street_sheet
$paiCol = [int]$manifest.pai_col
$jieCol = [int]$manifest.jie_col
$header1 = [string]$manifest.header1
$header2 = [string]$manifest.header2
$orangeRows = @($manifest.orange_rows | ForEach-Object { [int]$_ })

if (-not (Test-Path -LiteralPath $path)) {
    Write-Output "ERROR,file_not_found,$path"
    exit 1
}
if ($orangeRows.Count -eq 0) {
    Write-Output "ERROR,no_orange_rows"
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
    $oldCalc = $Excel.Calculation
    $Excel.Calculation = $xlCalculationManual

    $wb = $Excel.Workbooks.Open($path, 0, $false, 5, "", "", $true, "", "", $false, $false, $false)
    $swOpen.Stop()
    Write-Output "TIMING,open_ms,$($swOpen.ElapsedMilliseconds)"

    if ($wb.Worksheets.Count -le $sheetIndex) {
        Write-Output "ERROR,no_sheet"
        exit 1
    }

    $src = $wb.Worksheets.Item($sheetIndex + 1)
    $sourceName = $src.Name

    $outCol1 = Find-HeaderCol $src $headerRow $header1
    $outCol2 = Find-HeaderCol $src $headerRow $header2
    if ($outCol1 -gt 0 -and $outCol2 -gt 0) {
        # reuse
    }
    elseif ($outCol1 -eq 0 -and $outCol2 -eq 0) {
        $last = Get-HeaderLastCol $src $headerRow
        $outCol1 = $last + 1
        $outCol2 = $last + 2
        $src.Cells.Item($headerRow, $outCol1).Value2 = $header1
        $src.Cells.Item($headerRow, $outCol2).Value2 = $header2
    }
    else {
        Write-Output "ERROR,partial_headers"
        exit 1
    }

    $lastCol = Get-HeaderLastCol $src $headerRow

    $swFormula = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($row in $orangeRows) {
        Set-SplitFormulas $src $row $sourceCol $paiCol $jieCol $outCol1 $outCol2
    }
    $swFormula.Stop()
    Write-Output "TIMING,formula_ms,$($swFormula.ElapsedMilliseconds)"

    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.Name -eq $streetSheet) {
            $ws.Delete()
            break
        }
    }

    $dest = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $dest.Name = $streetSheet

    $swHeader = [System.Diagnostics.Stopwatch]::StartNew()
    Copy-HeaderLayout $Excel $src $dest $headerRow $lastCol
    $swHeader.Stop()
    Write-Output "TIMING,header_ms,$($swHeader.ElapsedMilliseconds)"

    $swCopy = [System.Diagnostics.Stopwatch]::StartNew()
    $destRow = $dataStartRow
    $pasted = 0
    foreach ($row in $orangeRows) {
        $srcRange = $src.Range(
            $src.Cells.Item($row, 1),
            $src.Cells.Item($row, $lastCol)
        )
        $dstRange = $dest.Range(
            $dest.Cells.Item($destRow, 1),
            $dest.Cells.Item($destRow, $lastCol)
        )
        $srcRange.Copy($dstRange) | Out-Null

        $streetVal = Get-StreetGValue $src $row $sourceCol $paiCol $jieCol
        if ($null -ne $streetVal) {
            $dest.Cells.Item($destRow, $sourceCol).Value2 = $streetVal
        }
        $dest.Cells.Item($destRow, $sourceCol).NumberFormat = $src.Cells.Item($row, $sourceCol).NumberFormat

        $destRow++
        $pasted++
    }
    Clear-CopyMode $Excel
    $swCopy.Stop()
    Write-Output "TIMING,copy_ms,$($swCopy.ElapsedMilliseconds)"

    $swSave = [System.Diagnostics.Stopwatch]::StartNew()
    Save-WorkbookForce $wb $path
    $swSave.Stop()
    Write-Output "TIMING,save_ms,$($swSave.ElapsedMilliseconds)"

    $saved = $true
    Write-Output "SAVED,$path"
    Write-Output "STATS,$sourceName,$outCol1,$outCol2,$($orangeRows.Count),$pasted"
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
