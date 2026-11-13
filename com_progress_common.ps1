# COM 脚本共用：带 flush 的进度输出，供 Python 实时读取

# 必须用 Console.Out 而非 Write-Output，否则进度行会污染函数 return 值

function Write-ProgressLine([string]$Message) {

    [Console]::Out.WriteLine($Message)

    try { [Console]::Out.Flush() } catch { }

}



function Get-ProgressStep([int]$Total, [int]$Parts = 20) {

    if ($Total -le 0) { return 1 }

    return [Math]::Max(1, [int][Math]::Ceiling($Total / [double]$Parts))

}

