param([Parameter(Mandatory = $true)][string]$ManifestPath)

$ErrorActionPreference = "Stop"
$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$startRow = [int]$manifest.startRow
$firstOnly = [bool]$manifest.firstOnly
$missing = [Type]::Missing

function Get-LastCell($Sheet, [int]$SearchOrder) {
    $Sheet.Cells.Find("*", $missing, -4123, 2, $SearchOrder, 2, $false, $false, $false)
}

$excel = New-Object -ComObject Excel.Application
$targetBook = $null
try {
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $targetBook = $excel.Workbooks.Open($manifest.files[0], 0, $true)
    if ($targetBook.Worksheets.Count -eq 0) { throw "The first workbook has no worksheets." }
    $totalRows = 0
    $usedFiles = 1
    $destinations = @{}
    $lastRows = @{}
    foreach ($sheet in $targetBook.Worksheets) {
        $destinations[$sheet.Name] = $sheet
        $lastRows[$sheet.Name] = [int](Get-LastCell $sheet 1).Row
        $totalRows += [Math]::Max(0, $lastRows[$sheet.Name] - $startRow + 1)
        if ($firstOnly) { break }
    }

    for ($index = 1; $index -lt $manifest.files.Count; $index++) {
        $sourceBook = $excel.Workbooks.Open($manifest.files[$index], 0, $true)
        try {
            if ($sourceBook.Worksheets.Count -eq 0) { continue }
            foreach ($source in $sourceBook.Worksheets) {
                $lastRow = [int](Get-LastCell $source 1).Row
                if ($firstOnly) {
                    $key = $targetBook.Worksheets.Item(1).Name
                } else {
                    $key = $source.Name
                }
                if (-not $destinations.ContainsKey($key)) {
                    $source.Copy($missing, $targetBook.Sheets.Item($targetBook.Sheets.Count))
                    $destinations[$key] = $targetBook.Worksheets.Item($key)
                    $lastRows[$key] = $lastRow
                } elseif ($lastRow -ge $startRow) {
                    $target = $destinations[$key]
                    $destRow = [Math]::Max($lastRows[$key] + 1, $startRow)
                    $destLast = $destRow + $lastRow - $startRow
                    if ($destLast -gt 1048576) { throw "Excel row limit exceeded: $key" }
                    $lastCol = [int](Get-LastCell $source 2).Column
                    $range = $source.Range($source.Cells.Item($startRow, 1), $source.Cells.Item($lastRow, $lastCol))
                    $range.Copy($target.Cells.Item($destRow, 1)) | Out-Null
                    $lastRows[$key] = $destLast
                }
                $totalRows += [Math]::Max(0, $lastRow - $startRow + 1)
                if ($firstOnly) { break }
            }
            $usedFiles++
        } finally {
            $sourceBook.Close($false)
        }
    }
    $targetBook.SaveAs($manifest.output, 51)
    @{rows = $totalRows; files = $usedFiles} | ConvertTo-Json |
        Set-Content -LiteralPath $manifest.result -Encoding UTF8
} finally {
    if ($null -ne $targetBook) { $targetBook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
