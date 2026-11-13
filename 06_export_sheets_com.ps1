param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Release-ComObjectSafe($obj) {
    if ($null -ne $obj) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj)
    }
}

function Get-DesiredOrder([string]$Primary, [string[]]$SourceSheets) {
    $rest = @($SourceSheets | Where-Object { $_ -ne $Primary })
    return @($Primary) + @($rest)
}

function Find-SourceSheet($SourceWorkbook, [string]$Name) {
    foreach ($ws in @($SourceWorkbook.Worksheets)) {
        if ($ws.Name -eq $Name) {
            return $ws
        }
    }
    return $null
}

function Build-TargetWorkbook(
    $Excel,
    $SourceWorkbook,
    [string[]]$DesiredOrder,
    [string[]]$SourceSheets,
    [string]$DestPath,
    [string]$Mode
) {
    $newWb = $Excel.Workbooks.Add()
    $placeholder = $newWb.Worksheets.Item(1)

    for ($i = 0; $i -lt $DesiredOrder.Count; $i++) {
        $name = $DesiredOrder[$i]
        $srcSheet = Find-SourceSheet $SourceWorkbook $name
        if ($null -eq $srcSheet) {
            throw "Source sheet not found: $name"
        }

        if ($i -eq 0) {
            $srcSheet.Copy($newWb.Worksheets(1)) | Out-Null
            $placeholder.Delete() | Out-Null
            $placeholder = $null
        }
        else {
            $lastIndex = $newWb.Worksheets.Count
            $srcSheet.Copy([Type]::Missing, $newWb.Worksheets($lastIndex)) | Out-Null
        }
    }

    if ($Mode -eq "overwrite" -and (Test-Path -LiteralPath $DestPath)) {
        $oldWb = $Excel.Workbooks.Open($DestPath, $null, $true)
        try {
            foreach ($ws in @($oldWb.Worksheets)) {
                if ($SourceSheets -notcontains $ws.Name) {
                    $lastIndex = $newWb.Worksheets.Count
                    $ws.Copy([Type]::Missing, $newWb.Worksheets($lastIndex)) | Out-Null
                }
            }
        }
        finally {
            $oldWb.Close($false) | Out-Null
            Release-ComObjectSafe $oldWb
        }
    }

    $destDir = Split-Path -Parent $DestPath
    if (-not (Test-Path -LiteralPath $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }
    if (Test-Path -LiteralPath $DestPath) {
        Remove-Item -LiteralPath $DestPath -Force
    }

    $newWb.SaveAs($DestPath, 51)
    $newWb.Close($false) | Out-Null
    Release-ComObjectSafe $newWb
}

$excel = $null
$srcWb = $null

try {
    if (-not (Test-Path -LiteralPath $ManifestPath)) {
        Write-Output "ERROR,manifest_not_found,$ManifestPath"
        exit 1
    }

    $manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $sourcePath = [string]$manifest.source
    $sourceSheets = @($manifest.source_sheets | ForEach-Object { [string]$_ })
    $targets = @($manifest.targets)

    if (-not (Test-Path -LiteralPath $sourcePath)) {
        Write-Output "ERROR,source_not_found,$sourcePath"
        exit 1
    }
    if ($sourceSheets.Count -eq 0 -or $targets.Count -eq 0) {
        Write-Output "ERROR,empty_manifest"
        exit 1
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $srcWb = $excel.Workbooks.Open($sourcePath, $null, $true)

    $count = 0
    $outputDir = Split-Path -Parent ([string]$targets[0].dest)

    foreach ($tgt in $targets) {
        $primary = [string]$tgt.primary
        $dest = [string]$tgt.dest
        $mode = [string]$tgt.mode
        $desiredOrder = Get-DesiredOrder $primary $sourceSheets

        Build-TargetWorkbook $excel $srcWb $desiredOrder $sourceSheets $dest $mode

        if ($mode -eq "create") {
            Write-Output "CREATED,$primary,$dest"
        }
        else {
            Write-Output "SYNCED,$primary,$dest"
        }
        $count++
    }

    Write-Output "STATS,$count,$outputDir"
    exit 0
}
catch {
    Write-Output "ERROR,$($_.Exception.Message)"
    exit 1
}
finally {
    if ($null -ne $srcWb) {
        $srcWb.Close($false) | Out-Null
        Release-ComObjectSafe $srcWb
    }
    if ($null -ne $excel) {
        $excel.Quit() | Out-Null
        Release-ComObjectSafe $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
