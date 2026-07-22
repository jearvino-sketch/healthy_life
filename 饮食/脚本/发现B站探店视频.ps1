[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Term,
    [Parameter(Mandatory = $true)]
    [string]$City,
    [Parameter(Mandatory = $true)]
    [string]$Area,
    [string]$DiscoveryJson = (Join-Path ([IO.Path]::GetTempPath()) 'healthy_life_bili_discovery.json'),
    [ValidateRange(250, 5000)]
    [int]$MinDelayMs = 400
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$headers = @{ Referer = 'https://search.bilibili.com/'; 'User-Agent' = 'Mozilla/5.0' }
$encodedTerm = [uri]::EscapeDataString($Term)

function Invoke-SearchPage([int]$Page) {
    Start-Sleep -Milliseconds $MinDelayMs
    $uri = "https://api.bilibili.com/x/web-interface/search/type?search_type=video&keyword=$encodedTerm&page=$Page"
    $response = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 30
    if ([int]$response.code -ne 0) { throw "Bilibili search API code $($response.code): $($response.message)" }
    return $response.data
}

$first = $null
try {
    $first = Invoke-SearchPage -Page 1
}
catch {
    if ($_.Exception.Message -match '\b412\b') {
        $oldRows = @()
        $oldRuns = @()
        if (Test-Path -LiteralPath $DiscoveryJson) {
            $oldDiscovery = Get-Content -LiteralPath $DiscoveryJson -Raw -Encoding UTF8 | ConvertFrom-Json
            $oldRows = @($oldDiscovery.uniqueRows)
            $oldRuns = @($oldDiscovery.runs | Where-Object { -not (([string]$_.city -eq $City) -and ([string]$_.area -eq $Area) -and ([string]$_.term -eq $Term)) })
        }
        [pscustomobject]@{
            uniqueRows = $oldRows
            runs = @($oldRuns + [pscustomobject]@{
                city = $City; area = $Area; term = $Term
                lastPageAdvertised = 0; lastPageFetched = 0; rawRows = 0
                failures = @([pscustomobject]@{ page = 1; reason = $_.Exception.Message })
                stopReason = 'HTTP 412; stopped without retry'
            })
        } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $DiscoveryJson -Encoding UTF8
    }
    throw
}
$advertisedLastPage = [int]$first.numPages
if ($advertisedLastPage -lt 1) { throw 'Bilibili search returned no valid page count.' }

$fetched = New-Object 'System.Collections.Generic.List[object]'
$failures = New-Object 'System.Collections.Generic.List[object]'
$actualLastPage = 0
for ($page = 1; $page -le $advertisedLastPage; $page++) {
    try {
        $data = if ($page -eq 1) { $first } else { Invoke-SearchPage -Page $page }
        $results = @($data.result | Where-Object { [string]$_.type -eq 'video' })
        if (-not $results.Count) { break }
        foreach ($item in $results) {
            [void]$fetched.Add([pscustomobject]@{
                city = $City; area = $Area; term = $Term; page = $page
                bvid = [string]$item.bvid
                title = (([string]$item.title) -replace '<[^>]+>', '')
                author = [string]$item.author; mid = [string]$item.mid
                pubdate = [long]$item.pubdate; pubstr = ''
                play = [long]$item.play; comments = [long]$item.review; danmaku = [long]$item.danmaku
                duration = [string]$item.duration; tag = [string]$item.tag
                description = (([string]$item.description) -replace '<[^>]+>', '')
                sourceUrl = "https://search.bilibili.com/all?keyword=$encodedTerm&page=$page"
                areaMatch = $true; foodMatch = $true; commentPass = ([long]$item.review -gt 20); preliminaryPass = $true
                queries = @($Term); areas = @($Area); pages = @($page)
            })
        }
        $actualLastPage = $page
    }
    catch {
        [void]$failures.Add([pscustomobject]@{ page = $page; reason = $_.Exception.Message })
        if ($_.Exception.Message -match '\b412\b') { throw }
        break
    }
}

$existingRows = @()
$existingRuns = @()
if (Test-Path -LiteralPath $DiscoveryJson) {
    $existing = Get-Content -LiteralPath $DiscoveryJson -Raw -Encoding UTF8 | ConvertFrom-Json
    $existingRows = @($existing.uniqueRows | Where-Object { -not (([string]$_.city -eq $City) -and ([string]$_.area -eq $Area) -and (@([string]$_.queries) -contains $Term)) })
    $existingRuns = @($existing.runs | Where-Object { -not (([string]$_.city -eq $City) -and ([string]$_.area -eq $Area) -and ([string]$_.term -eq $Term)) })
}

$rowsByBvid = @{}
foreach ($row in @($existingRows) + @($fetched)) {
    $bvid = [string]$row.bvid
    if ([string]::IsNullOrWhiteSpace($bvid)) { continue }
    if (-not $rowsByBvid.ContainsKey($bvid)) {
        $rowsByBvid[$bvid] = $row
        continue
    }
    $saved = $rowsByBvid[$bvid]
    $saved.queries = @($saved.queries + $row.queries | Sort-Object -Unique)
    $saved.areas = @($saved.areas + $row.areas | Sort-Object -Unique)
    $saved.pages = @($saved.pages + $row.pages | Sort-Object -Unique)
}

$run = [pscustomobject]@{
    city = $City; area = $Area; term = $Term
    lastPageAdvertised = $advertisedLastPage; lastPageFetched = $actualLastPage
    rawRows = $fetched.Count; failures = @($failures)
    stopReason = if ($failures.Count) { 'Stopped after API error' } elseif ($actualLastPage -lt $advertisedLastPage) { 'Search results exhausted' } else { 'Fetched advertised final page' }
}
[pscustomobject]@{
    uniqueRows = @($rowsByBvid.Values | Sort-Object bvid)
    runs = @($existingRuns + $run)
} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $DiscoveryJson -Encoding UTF8

[pscustomobject]@{
    Term = $Term; Target = "$City/$Area"; Pages = $actualLastPage
    Cards = $fetched.Count; UniqueBvids = $rowsByBvid.Count; Output = $DiscoveryJson
} | Format-List
