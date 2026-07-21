[CmdletBinding()]
param(
    [string]$SourceCsv = '',
    [string]$CollectorScript = '',
    [string]$OutputDirectory = (Join-Path ([IO.Path]::GetTempPath()) 'healthy_life_bili_comments'),
    [string]$StatusCsv = '',
    [ValidateRange(250, 5000)]
    [int]$MinDelayMs = 500,
    [ValidateRange(0, 1000000)]
    [int]$BatchSize = 0,
    [switch]$RetryFailed,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$dietRoot = Split-Path -Parent $PSScriptRoot
$dataRoot = Join-Path $dietRoot '饭店数据'
if ([string]::IsNullOrWhiteSpace($SourceCsv)) {
    $SourceCsv = Join-Path $dataRoot '视频来源.csv'
}
if ([string]::IsNullOrWhiteSpace($CollectorScript)) {
    $CollectorScript = Join-Path $PSScriptRoot '采集B站评论.ps1'
}
if ([string]::IsNullOrWhiteSpace($StatusCsv)) {
    $StatusCsv = Join-Path $dataRoot '评论采集状态.csv'
}

foreach ($requiredPath in @($SourceCsv, $CollectorScript)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required file not found: $requiredPath"
    }
}
$null = New-Item -ItemType Directory -Force -Path $OutputDirectory

$sources = @(
    Import-Csv -LiteralPath $SourceCsv -Encoding UTF8 |
        Sort-Object @{ Expression = { [long]$_.评论数 } }, BV号
)
if ($BatchSize -gt 0) {
    $sources = @($sources | Select-Object -First $BatchSize)
}
$statusByBvid = @{}
if (Test-Path -LiteralPath $StatusCsv) {
    foreach ($row in @(Import-Csv -LiteralPath $StatusCsv -Encoding UTF8)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$row.BV号)) {
            $statusByBvid[[string]$row.BV号] = $row
        }
    }
}

$previous412 = @($statusByBvid.Values | Where-Object { [string]$_.错误 -match '\b412\b' })
if ($previous412.Count -gt 0 -and -not $RetryFailed) {
    $failedBvids = (($previous412 | ForEach-Object { [string]$_.BV号 }) -join '、')
    throw "状态表存在 HTTP 412 失败（$failedBvids），为避免自动重试已停止。确认站点限制解除后，显式传入 -RetryFailed。"
}

function Save-Status {
    @($statusByBvid.Values | Sort-Object 城市, 区县或县级市, BV号) |
        Export-Csv -LiteralPath $StatusCsv -NoTypeInformation -Encoding UTF8
}

function New-StatusRowFromFile {
    param(
        [Parameter(Mandatory = $true)]$Source,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $complete = [bool]$raw.完整
    return [pscustomobject]@{
        BV号           = [string]$Source.BV号
        城市           = [string]$Source.城市
        区县或县级市   = [string]$Source.区县或县级市
        页面评论总数   = [int]$raw.页面评论总数
        实际抓取数     = [int]$raw.实际抓取数
        顶层评论数     = [int]$raw.顶层评论数
        回复数         = [int]$raw.回复数
        完整           = $complete
        采集状态       = if ($complete) { "完整（$($raw.实际抓取数)/$($raw.页面评论总数)，含回复）" } else { "未完整（$($raw.实际抓取数)/$($raw.页面评论总数)，含回复）" }
        采集时间       = [string]$raw.采集时间
        输出文件       = $Path
        错误           = ''
    }
}

$completed = 0
$skipped = 0
$failed = 0
foreach ($source in $sources) {
    $bvid = [string]$source.BV号
    $outputPath = Join-Path $OutputDirectory "$bvid.json"

    if (-not $RetryFailed -and $statusByBvid.ContainsKey($bvid) -and [string]$statusByBvid[$bvid].采集状态 -eq '失败') {
        $skipped++
        continue
    }

    if (-not $Force -and (Test-Path -LiteralPath $outputPath)) {
        try {
            $statusByBvid[$bvid] = New-StatusRowFromFile -Source $source -Path $outputPath
            $skipped++
            Save-Status
            continue
        }
        catch {
            # A malformed checkpoint is re-fetched below.
        }
    }

    try {
        $resultLines = @(& $CollectorScript -Bvid $bvid -OutputDirectory $OutputDirectory -MinDelayMs $MinDelayMs)
        $resultLine = [string]($resultLines | Select-Object -Last 1)
        $result = $resultLine | ConvertFrom-Json
        if (-not [string]::IsNullOrWhiteSpace([string]$result.错误)) {
            throw [string]$result.错误
        }
        $statusByBvid[$bvid] = New-StatusRowFromFile -Source $source -Path ([string]$result.输出文件)
        $completed++
    }
    catch {
        $message = $_.Exception.Message
        $statusByBvid[$bvid] = [pscustomobject]@{
            BV号           = $bvid
            城市           = [string]$source.城市
            区县或县级市   = [string]$source.区县或县级市
            页面评论总数   = ''
            实际抓取数     = ''
            顶层评论数     = ''
            回复数         = ''
            完整           = $false
            采集状态       = '失败'
            采集时间       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            输出文件       = ''
            错误           = $message
        }
        $failed++
        Save-Status
        if ($message -match '\b412\b') {
            throw
        }
        continue
    }

    Save-Status
    $processed = $completed + $skipped + $failed
    if ($processed % 10 -eq 0 -or $processed -eq $sources.Count) {
        [pscustomobject]@{
            已处理 = $processed
            总数   = $sources.Count
            新抓取 = $completed
            复用断点 = $skipped
            失败   = $failed
            BV号   = $bvid
        } | ConvertTo-Json -Compress
    }
}

[pscustomobject]@{
    来源视频数 = $sources.Count
    新抓取 = $completed
    复用断点 = $skipped
    失败 = $failed
    状态表 = $StatusCsv
    原始目录 = $OutputDirectory
} | ConvertTo-Json -Compress
