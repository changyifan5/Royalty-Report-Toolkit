# Excel B 列 AutoFilter 共用：扫描 + 整块复制（09 / 11 等）

$script:xlUp = -4162
$script:xlCellTypeVisible = 12
$script:xlFilterFontColor = 9
$script:xlFilterCellColor = 8
$script:xlCalculationAutomatic = -4105

function Get-ExcelRgbLong([int]$R, [int]$G, [int]$B) {
    return [int]($R + ($G * 256) + ($B * 65536))
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

function Get-LastDataRow($Worksheet, [int]$DataStartRow, [int]$CheckCol = 1) {
    $last = [int]$Worksheet.Cells.Item($Worksheet.Rows.Count, $CheckCol).End($script:xlUp).Row
    if ($last -lt $DataStartRow) { return $DataStartRow - 1 }
    return $last
}

function Get-FilterColorCalibrationEntry([string]$FilterMode) {
    $defaults = @{
        'red_font'    = @{ rgb = @(255, 0, 0); tolerance = 12; theme = $null; tint = 0.0; kind = 'font' }
        'blue_font'   = @{ rgb = @(75, 172, 198); tolerance = 12; theme = 8; tint = 0.0; kind = 'font' }
        'purple_font' = @{ rgb = @(112, 48, 160); tolerance = 12; theme = $null; tint = 0.0; kind = 'font' }
        'orange_font' = @{ rgb = @(192, 80, 77); tolerance = 12; theme = 5; tint = 0.0; kind = 'font' }
        'gray_bold'   = @{ rgb = @(217, 217, 217); tolerance = 20; theme = $null; tint = 0.0; kind = 'interior' }
    }
    if (-not $defaults.ContainsKey($FilterMode)) {
        return $null
    }
    $entry = @{
        rgb       = @($defaults[$FilterMode].rgb[0], $defaults[$FilterMode].rgb[1], $defaults[$FilterMode].rgb[2])
        tolerance = [int]$defaults[$FilterMode].tolerance
        theme     = $defaults[$FilterMode].theme
        tint      = [double]$defaults[$FilterMode].tint
        kind      = [string]$defaults[$FilterMode].kind
    }

    $allPath = Join-Path $PSScriptRoot "filter_color_calibration.json"
    if (Test-Path -LiteralPath $allPath) {
        try {
            $all = Get-Content -LiteralPath $allPath -Encoding UTF8 -Raw | ConvertFrom-Json
            if ($all.modes -and $all.modes.$FilterMode) {
                $mode = $all.modes.$FilterMode
                if ($mode.rgb -and @($mode.rgb).Count -ge 3) {
                    $entry.rgb = @([int]$mode.rgb[0], [int]$mode.rgb[1], [int]$mode.rgb[2])
                }
                if ($null -ne $mode.tolerance) { $entry.tolerance = [int]$mode.tolerance }
                if ($mode.PSObject.Properties.Name -contains 'excel_theme' -and $null -ne $mode.excel_theme) {
                    $entry.theme = [int]$mode.excel_theme
                }
                if ($null -ne $mode.excel_tint) { $entry.tint = [double]$mode.excel_tint }
                if ($mode.kind) { $entry.kind = [string]$mode.kind }
            }
        }
        catch { }
    }

    $modeFiles = @{
        'orange_font' = 'orange_font_calibration.json'
        'red_font'    = 'red_font_calibration.json'
        'blue_font'   = 'blue_font_calibration.json'
        'purple_font' = 'purple_font_calibration.json'
    }
    if ($modeFiles.ContainsKey($FilterMode)) {
        $calPath = Join-Path $PSScriptRoot $modeFiles[$FilterMode]
        if (Test-Path -LiteralPath $calPath) {
            try {
                $cal = Get-Content -LiteralPath $calPath -Encoding UTF8 -Raw | ConvertFrom-Json
                if ($cal.rgb -and @($cal.rgb).Count -ge 3) {
                    $entry.rgb = @([int]$cal.rgb[0], [int]$cal.rgb[1], [int]$cal.rgb[2])
                }
                if ($null -ne $cal.tolerance) { $entry.tolerance = [int]$cal.tolerance }
                if ($cal.PSObject.Properties.Name -contains 'excel_theme' -and $null -ne $cal.excel_theme) {
                    $entry.theme = [int]$cal.excel_theme
                }
                elseif ($cal.PSObject.Properties.Name -contains 'excel_theme' -and $null -eq $cal.excel_theme) {
                    $entry.theme = $null
                }
                if ($null -ne $cal.excel_tint) { $entry.tint = [double]$cal.excel_tint }
            }
            catch { }
        }
    }

    return $entry
}

function Get-CalibratedFilterColorLong([string]$FilterMode) {
    $entry = Get-FilterColorCalibrationEntry $FilterMode
    if ($null -eq $entry) { return 0 }
    return Get-ExcelRgbLong $entry.rgb[0] $entry.rgb[1] $entry.rgb[2]
}

function Get-OrangeCalibrationRef() {
    $entry = Get-FilterColorCalibrationEntry 'orange_font'
    if ($null -eq $entry) {
        return @(192, 80, 77), 12, 5, 0.0
    }
    $theme = 5
    if ($null -ne $entry.theme) { $theme = [int]$entry.theme }
    return @($entry.rgb[0], $entry.rgb[1], $entry.rgb[2]), [int]$entry.tolerance, $theme, [double]$entry.tint
}

function Get-CalibratedOrangeColorLong() {
    return Get-CalibratedFilterColorLong 'orange_font'
}

function Get-ColorRgbFromLong([int]$ColorLong) {
    return @(
        [int](($ColorLong -shr 16) -band 0xFF),
        [int](($ColorLong -shr 8) -band 0xFF),
        [int]($ColorLong -band 0xFF)
    )
}

function Test-FontColorMatchesOrange([int]$ColorLong) {
    if ($ColorLong -le 0) { return $false }
    $rgb = Get-ColorRgbFromLong $ColorLong
    $ref, $tol, $null, $null = Get-OrangeCalibrationRef
    return (
        [Math]::Abs($rgb[0] - $ref[0]) -le $tol -and
        [Math]::Abs($rgb[1] - $ref[1]) -le $tol -and
        [Math]::Abs($rgb[2] - $ref[2]) -le $tol
    )
}

function Test-FontColorMatchesRed([int]$ColorLong) {
    if ($ColorLong -le 0) { return $false }
    if (Test-FontColorMatchesOrange $ColorLong) { return $false }
    $r, $g, $b = Get-ColorRgbFromLong $ColorLong
    if ($r -lt 160) { return $false }
    if ($b -gt 110 -or $g -gt 95) { return $false }
    if (($r - $g) -lt 70 -or ($r - $b) -lt 70) { return $false }
    return ($r -gt $g -and $r -gt $b)
}

function Test-FontColorMatchesPurple([int]$ColorLong) {
    if ($ColorLong -le 0) { return $false }
    $r, $g, $b = Get-ColorRgbFromLong $ColorLong
    if ($r -ge 80 -and $b -ge 100 -and $g -le ([Math]::Min($r, $b) - 25)) { return $true }
    if ($b -ge 120 -and $r -ge 60 -and $g -le 95 -and $b -ge $r) { return $true }
    return $false
}

function Test-FontColorMatchesBlue([int]$ColorLong) {
    if ($ColorLong -le 0) { return $false }
    if (Test-FontColorMatchesPurple $ColorLong) { return $false }
    $r, $g, $b = Get-ColorRgbFromLong $ColorLong
    return ($b -ge 150 -and $r -le 130 -and $b -gt ($r + 30) -and $b -ge ($g + 20))
}

function Test-InteriorColorMatchesGray([int]$ColorLong) {
    if ($ColorLong -le 0) { return $false }
    $r, $g, $b = Get-ColorRgbFromLong $ColorLong
    if (([Math]::Max($r, [Math]::Max($g, $b)) - [Math]::Min($r, [Math]::Min($g, $b))) -gt 45) {
        return $false
    }
    $avg = ($r + $g + $b) / 3.0
    return ($avg -ge 70 -and $avg -le 210)
}

function Test-CellFontOrange($Cell) {
    if ($null -eq $Cell) { return $false }
    foreach ($getter in @(
            { [int]$Cell.DisplayFormat.Font.Color },
            { [int]$Cell.Font.Color }
        )) {
        try {
            if (Test-FontColorMatchesOrange (& $getter)) { return $true }
        }
        catch { }
    }
    try {
        $ref, $null, $theme, $tint = Get-OrangeCalibrationRef
        $font = $Cell.Font
        if ([int]$font.ThemeColor -eq $theme) {
            $cellTint = 0.0
            try { $cellTint = [double]$font.TintAndShade } catch { }
            if ([Math]::Abs($cellTint - $tint) -lt 0.01) { return $true }
        }
    }
    catch { }
    return $false
}

function Test-CellFontRed($Cell) {
    if ($null -eq $Cell) { return $false }
    foreach ($getter in @(
            { [int]$Cell.DisplayFormat.Font.Color },
            { [int]$Cell.Font.Color }
        )) {
        try {
            if (Test-FontColorMatchesRed (& $getter)) { return $true }
        }
        catch { }
    }
    return $false
}

function Test-CellFontBlue($Cell) {
    if ($null -eq $Cell) { return $false }
    foreach ($getter in @(
            { [int]$Cell.DisplayFormat.Font.Color },
            { [int]$Cell.Font.Color }
        )) {
        try {
            if (Test-FontColorMatchesBlue (& $getter)) { return $true }
        }
        catch { }
    }
    return $false
}

function Test-CellFontPurple($Cell) {
    if ($null -eq $Cell) { return $false }
    foreach ($getter in @(
            { [int]$Cell.DisplayFormat.Font.Color },
            { [int]$Cell.Font.Color }
        )) {
        try {
            if (Test-FontColorMatchesPurple (& $getter)) { return $true }
        }
        catch { }
    }
    return $false
}

function Test-CellGrayBold($Cell) {
    if ($null -eq $Cell) { return $false }
    try {
        if (-not $Cell.Font.Bold) { return $false }
    }
    catch {
        return $false
    }
    foreach ($getter in @(
            { [int]$Cell.DisplayFormat.Interior.Color },
            { [int]$Cell.Interior.Color }
        )) {
        try {
            if (Test-InteriorColorMatchesGray (& $getter)) { return $true }
        }
        catch { }
    }
    return $false
}

function Test-CellMatchesFilterMode($Cell, [string]$FilterMode) {
    switch ($FilterMode) {
        'orange_font' { return Test-CellFontOrange $Cell }
        'red_font'    { return Test-CellFontRed $Cell }
        'blue_font'   { return Test-CellFontBlue $Cell }
        'purple_font' { return Test-CellFontPurple $Cell }
        'gray_bold'   { return Test-CellGrayBold $Cell }
        default       { throw "unsupported_filter_mode:$FilterMode" }
    }
}

function Get-AutoFilterFontColorFromCell($Cell, $Worksheet, $Excel = $null) {
    if ($null -eq $Cell) { return 0 }
    $oldCalc = $null
    try {
        if ($null -ne $Worksheet) {
            $Worksheet.Calculate() | Out-Null
        }
        if ($null -ne $Excel) {
            $oldCalc = $Excel.Calculation
            $Excel.Calculation = $script:xlCalculationAutomatic
            if ($null -ne $Worksheet) {
                $Worksheet.Calculate() | Out-Null
            }
        }
        $display = 0
        try { $display = [int]$Cell.DisplayFormat.Font.Color } catch { }
        if ($display -gt 0) {
            return $display
        }
    }
    finally {
        if ($null -ne $Excel -and $null -ne $oldCalc) {
            try { $Excel.Calculation = $oldCalc } catch { }
        }
    }
    return Get-CellFilterFontColor $Cell
}

function Get-CellFilterFontColor($Cell) {
    if ($null -eq $Cell) { return 0 }
    foreach ($getter in @(
            { [int]$Cell.DisplayFormat.Font.Color },
            { [int]$Cell.Font.Color }
        )) {
        try {
            $color = & $getter
            if ($color -gt 0) { return [int]$color }
        }
        catch { }
    }
    return 0
}

function Get-CellFilterInteriorColor($Cell) {
    if ($null -eq $Cell) { return 0 }
    foreach ($getter in @(
            { [int]$Cell.Interior.Color },
            { [int]$Cell.DisplayFormat.Interior.Color }
        )) {
        try {
            $color = & $getter
            if ($color -gt 0) { return [int]$color }
        }
        catch { }
    }
    return 0
}

function Find-FirstMatchingRow(
    $Worksheet,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$FilterCol,
    [string]$FilterMode
) {
    for ($r = [int]$LastRow; $r -ge $DataStartRow; $r--) {
        $cell = $Worksheet.Cells.Item($r, $FilterCol)
        if (Test-CellMatchesFilterMode $cell $FilterMode) {
            return [int]$r
        }
    }
    return 0
}

function Find-FilterColorInColumn(
    $Worksheet,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$FilterCol,
    [string]$FilterMode,
    [scriptblock]$OnProgress = $null
) {
    $total = [int]($LastRow - $DataStartRow + 1)
    $step = [Math]::Max(1, [int][Math]::Ceiling($total / 20.0))
    $scanned = 0
    for ($r = [int]$LastRow; $r -ge $DataStartRow; $r--) {
        $scanned++
        if ($null -ne $OnProgress -and ($scanned -eq 1 -or ($scanned % $step) -eq 0)) {
            & $OnProgress $scanned $total | Out-Null
        }
        $cell = $Worksheet.Cells.Item($r, $FilterCol)
        if (-not (Test-CellMatchesFilterMode $cell $FilterMode)) {
            continue
        }
        if ($FilterMode -eq 'gray_bold') {
            $color = Get-CellFilterInteriorColor $cell
            if ($color -gt 0) {
                return $color
            }
            continue
        }
        $color = Get-CellFilterFontColor $cell
        if ($color -gt 0) {
            return $color
        }
        $calibrated = Get-CalibratedFilterColorLong $FilterMode
        if ($calibrated -gt 0) {
            return $calibrated
        }
    }
    $fallback = Get-CalibratedFilterColorLong $FilterMode
    if ($fallback -gt 0) {
        return $fallback
    }
    return 0
}

function Get-RowsByColumnProbe(
    $Worksheet,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$FilterCol,
    [string]$FilterMode,
    [scriptblock]$OnProgress = $null
) {
    $rows = New-Object System.Collections.Generic.List[int]
    $total = [int]($LastRow - $DataStartRow + 1)
    if ($total -le 0) {
        return @()
    }
    $step = [Math]::Max(1, [int][Math]::Ceiling($total / 20.0))

    if ($total -eq 1) {
        $cell = $Worksheet.Cells.Item($DataStartRow, $FilterCol)
        if (Test-CellMatchesFilterMode $cell $FilterMode) {
            $rows.Add($DataStartRow) | Out-Null
        }
        if ($null -ne $OnProgress) {
            & $OnProgress 1 1 | Out-Null
        }
        return @($rows.ToArray())
    }

    $rng = $Worksheet.Range(
        $Worksheet.Cells.Item($DataStartRow, $FilterCol),
        $Worksheet.Cells.Item($LastRow, $FilterCol)
    )
    $values = $rng.Value2

    for ($i = 1; $i -le $total; $i++) {
        if ($null -ne $OnProgress -and ($i -eq 1 -or ($i % $step) -eq 0 -or $i -eq $total)) {
            & $OnProgress $i $total | Out-Null
        }
        $v = $values[$i, 1]
        if ($null -eq $v) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$v)) { continue }
        $rowNum = [int]($DataStartRow + $i - 1)
        $cell = $Worksheet.Cells.Item($rowNum, $FilterCol)
        if (Test-CellMatchesFilterMode $cell $FilterMode) {
            $rows.Add($rowNum) | Out-Null
        }
    }
    return @($rows.ToArray())
}

function Get-VisibleFilterRowNumbers(
    $Worksheet,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$FilterCol
) {
    $rows = New-Object System.Collections.Generic.List[int]
    $colRange = $Worksheet.Range(
        $Worksheet.Cells.Item($DataStartRow, $FilterCol),
        $Worksheet.Cells.Item($LastRow, $FilterCol)
    )
    try {
        $visible = $colRange.SpecialCells($script:xlCellTypeVisible)
    }
    catch {
        return @()
    }
    foreach ($area in @($visible.Areas)) {
        $startRow = [int]$area.Row
        $count = [int]$area.Rows.Count
        for ($i = 0; $i -lt $count; $i++) {
            $rowNum = [int]($startRow + $i)
            if ($rowNum -ge $DataStartRow) {
                $rows.Add($rowNum) | Out-Null
            }
        }
    }
    return @($rows.ToArray())
}

function Apply-ColumnFilter(
    $Worksheet,
    [int]$HeaderRow,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$DataLastCol,
    [int]$FilterCol,
    [string]$FilterMode,
    [int]$FilterRefRow = 0,
    [int]$ColorLong = 0
) {
    if ($DataLastCol -lt 1) { $DataLastCol = 1 }
    if ($Worksheet.AutoFilterMode) {
        $Worksheet.ShowAllData() | Out-Null
    }
    $filterRange = $Worksheet.Range(
        $Worksheet.Cells.Item($HeaderRow, 1),
        $Worksheet.Cells.Item($LastRow, $DataLastCol)
    )
    $filterRange.AutoFilter() | Out-Null
    if ($ColorLong -le 0 -and $FilterRefRow -ge 1) {
        $cell = $Worksheet.Cells.Item($FilterRefRow, $FilterCol)
        if ($FilterMode -eq 'gray_bold') {
            $ColorLong = Get-CellFilterInteriorColor $cell
        }
        else {
            $ColorLong = Get-CellFilterFontColor $cell
        }
    }
    if ($ColorLong -le 0) {
        $ColorLong = Get-CalibratedFilterColorLong $FilterMode
    }
    if ($ColorLong -le 0) {
        $ColorLong = Find-FilterColorInColumn $Worksheet $DataStartRow $LastRow $FilterCol $FilterMode
    }
    if ($ColorLong -le 0) {
        throw "filter_color_not_found:$FilterMode"
    }
    $operator = if ($FilterMode -eq 'gray_bold') { $script:xlFilterCellColor } else { $script:xlFilterFontColor }
    $filterRange.AutoFilter($FilterCol, $ColorLong, $operator) | Out-Null
    return [int]$ColorLong
}

function Clear-ColumnFilter($Worksheet) {
    if ($null -eq $Worksheet) { return }
    if ($Worksheet.AutoFilterMode) {
        $Worksheet.ShowAllData() | Out-Null
        $Worksheet.AutoFilterMode = $false
    }
}

function Copy-FilteredDataToSheet(
    $Excel,
    $SourceSheet,
    $DestSheet,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$DataLastCol,
    [int]$DestStartRow
) {
    if ($DataLastCol -lt 1) { $DataLastCol = 1 }
    $srcData = $SourceSheet.Range(
        $SourceSheet.Cells.Item($DataStartRow, 1),
        $SourceSheet.Cells.Item($LastRow, $DataLastCol)
    )
    $visible = $srcData.SpecialCells($script:xlCellTypeVisible)
    $destCell = $DestSheet.Cells.Item($DestStartRow, 1)
    $visible.Copy($destCell) | Out-Null
    if ($null -ne $Excel) {
        $Excel.CutCopyMode = $false
    }
}

function Test-FilterVisibleMatch(
    $Worksheet,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$FilterCol,
    [int]$ExpectedCount
) {
    $visible = @(Get-VisibleFilterRowNumbers $Worksheet $DataStartRow $LastRow $FilterCol)
    $count = $visible.Count
    if ($count -le 0) { return $false, $count }
    if ($ExpectedCount -le 0) { return $true, $count }
    if ($ExpectedCount -le 10) {
        $tol = 0
    }
    else {
        $tol = [Math]::Max(3, [int][Math]::Ceiling($ExpectedCount * 0.02))
    }
    return ([Math]::Abs($count - $ExpectedCount) -le $tol), $count
}

function Enable-SplitColumnFilter(
    $Worksheet,
    $Excel,
    [int]$HeaderRow,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$DataLastCol,
    [int]$FilterCol,
    [string]$FilterMode,
    [int]$FilterRefRow,
    [int]$ExpectedCount
) {
    if ($DataLastCol -lt 1) { $DataLastCol = 1 }
    $colors = New-Object System.Collections.Generic.List[int]
    if ($FilterRefRow -ge 1) {
        $refCell = $Worksheet.Cells.Item($FilterRefRow, $FilterCol)
        if ($FilterMode -eq 'gray_bold') {
            $refColor = [int](Get-CellFilterInteriorColor $refCell)
        }
        else {
            $refColor = [int](Get-AutoFilterFontColorFromCell $refCell $Worksheet $Excel)
        }
        if ($refColor -gt 0) {
            [void]$colors.Add($refColor)
        }
    }
    $calibrated = [int](Get-CalibratedFilterColorLong $FilterMode)
    if ($calibrated -gt 0 -and -not $colors.Contains($calibrated)) {
        [void]$colors.Add($calibrated)
    }

    foreach ($color in @($colors)) {
        Clear-ColumnFilter $Worksheet
        try {
            Apply-ColumnFilter $Worksheet $HeaderRow $DataStartRow $LastRow $DataLastCol $FilterCol $FilterMode $FilterRefRow $color | Out-Null
        }
        catch {
            continue
        }
        $ok, $visibleCount = Test-FilterVisibleMatch $Worksheet $DataStartRow $LastRow $FilterCol $ExpectedCount
        if ($ok) {
            return @{ Active = $true; Color = [int]$color; Visible = [int]$visibleCount; Mode = 'filter' }
        }
    }

    Clear-ColumnFilter $Worksheet
    return @{ Active = $false; Color = 0; Visible = 0; Mode = 'data_rows' }
}

function Copy-DataRowsList(
    $Excel,
    $SourceSheet,
    $DestSheet,
    $DataRows,
    [int]$LastCol,
    [int]$DestStartRow
) {
    if ($LastCol -lt 1) { $LastCol = 1 }
    $rows = @($DataRows | ForEach-Object { [int]$_ })
    if ($rows.Count -eq 0) {
        return 0
    }
    $destRow = $DestStartRow
    $total = $rows.Count
    $idx = 0
    $step = [Math]::Max(1, [int][Math]::Ceiling($total / 20.0))
    foreach ($rowNum in $rows) {
        $idx++
        $srcRange = $SourceSheet.Range(
            $SourceSheet.Cells.Item($rowNum, 1),
            $SourceSheet.Cells.Item($rowNum, $LastCol)
        )
        $dstRange = $DestSheet.Range(
            $DestSheet.Cells.Item($destRow, 1),
            $DestSheet.Cells.Item($destRow, $LastCol)
        )
        $srcRange.Copy($dstRange) | Out-Null
        $destRow++
        if ($idx -eq 1 -or ($idx % $step) -eq 0 -or $idx -eq $total) {
            Write-ProgressLine "PROGRESS,copy_rows,$idx,$total"
        }
    }
    if ($null -ne $Excel) {
        $Excel.CutCopyMode = $false
    }
    return [int]$total
}

function Invoke-ScanByColumnFilter(
    $Worksheet,
    [int]$HeaderRow,
    [int]$DataStartRow,
    [int]$DataEndRow,
    [int]$FilterCol,
    [string]$FilterMode,
    [scriptblock]$OnProgress = $null
) {
    $dataLastCol = Get-HeaderLastCol $Worksheet $HeaderRow
    $lastRow = [Math]::Min($DataEndRow, (Get-LastDataRow $Worksheet $DataStartRow))
    if ($lastRow -lt $DataStartRow) {
        return @(), 0, $dataLastCol, "none"
    }

    $filterActive = $false
    $rows = @()
    $filterColor = 0
    $scanMethod = "filter"
    try {
        $filterColor = Get-CalibratedFilterColorLong $FilterMode
        if ($filterColor -le 0) {
            $filterColor = Find-FilterColorInColumn $Worksheet $DataStartRow $lastRow $FilterCol $FilterMode $OnProgress
        }
        if ($filterColor -gt 0 -and $FilterMode -ne 'gray_bold') {
            try {
                Apply-ColumnFilter $Worksheet $HeaderRow $DataStartRow $lastRow $dataLastCol $FilterCol $FilterMode 0 $filterColor | Out-Null
                $filterActive = $true
                $rows = @(Get-VisibleFilterRowNumbers $Worksheet $DataStartRow $lastRow $FilterCol)
            }
            catch {
                $filterActive = $false
                $rows = @()
            }
        }
        if ($rows.Count -eq 0 -and $filterActive) {
            Clear-ColumnFilter $Worksheet
            $filterActive = $false
        }
        if ($rows.Count -eq 0 -and $FilterMode -ne 'gray_bold') {
            $refRow = Find-FirstMatchingRow $Worksheet $DataStartRow $lastRow $FilterCol $FilterMode
            if ($refRow -gt 0) {
                $altColor = Get-CellFilterFontColor ($Worksheet.Cells.Item($refRow, $FilterCol))
                if ($altColor -gt 0 -and $altColor -ne $filterColor) {
                    try {
                        $filterColor = Apply-ColumnFilter $Worksheet $HeaderRow $DataStartRow $lastRow $dataLastCol $FilterCol $FilterMode $refRow $altColor
                        $filterActive = $true
                        $rows = @(Get-VisibleFilterRowNumbers $Worksheet $DataStartRow $lastRow $FilterCol)
                    }
                    catch {
                        if ($filterActive) {
                            Clear-ColumnFilter $Worksheet
                            $filterActive = $false
                        }
                    }
                }
            }
        }
        if ($rows.Count -eq 0) {
            if ($FilterMode -eq 'gray_bold' -and $filterColor -gt 0) {
                try {
                    Apply-ColumnFilter $Worksheet $HeaderRow $DataStartRow $lastRow $dataLastCol $FilterCol $FilterMode 0 $filterColor | Out-Null
                    $filterActive = $true
                    $rows = @(Get-VisibleFilterRowNumbers $Worksheet $DataStartRow $lastRow $FilterCol)
                }
                catch {
                    if ($filterActive) {
                        Clear-ColumnFilter $Worksheet
                        $filterActive = $false
                    }
                }
            }
        }
        if ($rows.Count -eq 0) {
            if ($filterActive) {
                Clear-ColumnFilter $Worksheet
                $filterActive = $false
            }
            $scanMethod = "probe"
            $probeTotal = [int]($lastRow - $DataStartRow + 1)
            if ($null -ne $OnProgress) {
                & $OnProgress 0 $probeTotal | Out-Null
            }
            $probeProgress = $null
            if ($null -ne $OnProgress) {
                $probeProgress = {
                    param([int]$Done, [int]$Total)
                    & $OnProgress $Done $Total | Out-Null
                }
            }
            $rows = @(Get-RowsByColumnProbe $Worksheet $DataStartRow $lastRow $FilterCol $FilterMode $probeProgress)
        }
    }
    finally {
        if ($filterActive) {
            Clear-ColumnFilter $Worksheet
        }
    }
    return @($rows), [int]$filterColor, [int]$dataLastCol, [string]$scanMethod
}

# 11 兼容别名
function Apply-OrangeColumnFilter(
    $Worksheet,
    [int]$HeaderRow,
    [int]$DataStartRow,
    [int]$LastRow,
    [int]$DataLastCol,
    [int]$FilterCol,
    [int]$FilterRefRow = 0,
    [int]$OrangeColorLong = 0
) {
    return Apply-ColumnFilter $Worksheet $HeaderRow $DataStartRow $LastRow $DataLastCol $FilterCol 'orange_font' $FilterRefRow $OrangeColorLong
}

function Clear-OrangeColumnFilter($Worksheet) { Clear-ColumnFilter $Worksheet }
function Find-OrangeFilterColorInColumn($Worksheet, [int]$DataStartRow, [int]$LastRow, [int]$FilterCol) {
    return Find-FilterColorInColumn $Worksheet $DataStartRow $LastRow $FilterCol 'orange_font'
}
function Get-OrangeRowsByColumnProbe($Worksheet, [int]$DataStartRow, [int]$LastRow, [int]$FilterCol) {
    return Get-RowsByColumnProbe $Worksheet $DataStartRow $LastRow $FilterCol 'orange_font'
}
function Get-VisibleOrangeRowNumbers($Worksheet, [int]$DataStartRow, [int]$LastRow, [int]$FilterCol) {
    return Get-VisibleFilterRowNumbers $Worksheet $DataStartRow $LastRow $FilterCol
}
function Get-OrangeFilterColor($Worksheet, [int]$FilterRefRow, [int]$FilterCol) {
    if ($FilterRefRow -ge 1) {
        $color = Get-CellFilterFontColor ($Worksheet.Cells.Item($FilterRefRow, $FilterCol))
        if ($color -gt 0) { return $color }
    }
    return Get-CalibratedOrangeColorLong
}
