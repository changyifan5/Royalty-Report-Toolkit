param([Parameter(Mandatory = $true)][string]$ManifestPath)

$ErrorActionPreference = "Stop"
$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = New-Object -ComObject Excel.Application
$output = $null
try {
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $output = $excel.Workbooks.Add(-4167)
    $placeholder = $output.Worksheets.Item(1)
    foreach ($report in $manifest.reports) {
        $source = $excel.Workbooks.Open($report.path, 0, $true)
        try {
            $target = $output.Worksheets.Add($placeholder)
            $target.Name = $report.title
            $range = $source.Worksheets.Item(1).UsedRange
            $range.Copy() | Out-Null
            $target.Cells.Item($range.Row, $range.Column).PasteSpecial(12) | Out-Null
            $excel.CutCopyMode = 0
        } finally {
            $source.Close($false)
        }
    }
    $placeholder.Delete()
    $output.SaveAs($manifest.output, 51)
} finally {
    if ($null -ne $output) {
        $output.Close($false)
    }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
