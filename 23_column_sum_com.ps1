param([Parameter(Mandatory = $true)][string]$ManifestPath)

$ErrorActionPreference = "Stop"
$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$column = [int]$manifest.column
$start = [int]$manifest.start
$excel = New-Object -ComObject Excel.Application
$book = $null
try {
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($manifest.source, 0, $true)
    $excel.Calculate()
    $results = @(
        foreach ($sheet in $book.Worksheets) {
            $used = $sheet.UsedRange
            $end = $used.Row + $used.Rows.Count - 1
            $total = "0"
            if ($end -ge $start) {
                $range = $sheet.Range($sheet.Cells.Item($start, $column), $sheet.Cells.Item($end, $column))
                $formula = "SUM(" + $range.Address() + ")"
                $value = $sheet.Evaluate($formula)
                if ($value -is [double]) {
                    $total = $value.ToString("R", [Globalization.CultureInfo]::InvariantCulture)
                } else {
                    $total = $null
                }
            }
            @{ sheet = $sheet.Name; total = $total }
        }
    )
    ConvertTo-Json -InputObject $results | Set-Content -LiteralPath $manifest.result -Encoding UTF8
} finally {
    if ($null -ne $book) { $book.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
