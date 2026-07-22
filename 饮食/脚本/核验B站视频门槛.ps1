[CmdletBinding()]
param(
    [string]$DiscoveryJson = (Join-Path ([IO.Path]::GetTempPath()) 'healthy_life_bili_discovery.json'),
    [string]$FollowerCacheJson = (Join-Path ([IO.Path]::GetTempPath()) 'healthy_life_bili_followers.json'),
    [string]$CandidateCsv = '',
    [string]$SourceCsv = '',
    [string]$DiscoveryRunCsv = '',
    [string]$CommentStatusCsv = '',
    [string]$CommentDir = (Join-Path ([IO.Path]::GetTempPath()) 'healthy_life_bili_comments'),
    [ValidateRange(0, 1000000000)]
    [int]$MinFollowers = 5000,
    [ValidateRange(0, 1000000000)]
    [int]$MinComments = 20,
    [ValidateRange(250, 5000)]
    [int]$MinDelayMs = 400,
    [switch]$SkipFollowerFetch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$dietRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    Split-Path -Parent $PSScriptRoot
}
elseif (-not [string]::IsNullOrWhiteSpace($CandidateCsv)) {
    Split-Path -Parent (Split-Path -Parent $CandidateCsv)
}
else {
    Join-Path (Get-Location).Path '饮食'
}
$dataRoot = Join-Path $dietRoot '饭店数据'
if ([string]::IsNullOrWhiteSpace($CandidateCsv)) {
    $CandidateCsv = Join-Path $dataRoot '视频候选.csv'
}
if ([string]::IsNullOrWhiteSpace($SourceCsv)) {
    $SourceCsv = Join-Path $dataRoot '视频来源.csv'
}
if ([string]::IsNullOrWhiteSpace($DiscoveryRunCsv)) {
    $DiscoveryRunCsv = Join-Path $dataRoot '视频发现批次.csv'
}
if ([string]::IsNullOrWhiteSpace($CommentStatusCsv)) {
    $CommentStatusCsv = Join-Path $dataRoot '评论采集状态.csv'
}

if (-not (Test-Path -LiteralPath $DiscoveryJson)) {
    throw "Discovery JSON not found: $DiscoveryJson"
}

$discovery = Get-Content -LiteralPath $DiscoveryJson -Raw -Encoding UTF8 | ConvertFrom-Json
$uniqueRows = @($discovery.uniqueRows)
$runs = @($discovery.runs)
$today = (Get-Date).ToString('yyyy-MM-dd')
$requestHeaders = @{
    Referer = 'https://www.bilibili.com/'
    'User-Agent' = 'Mozilla/5.0'
}

$cache = @{}
if (Test-Path -LiteralPath $FollowerCacheJson) {
    $cacheObject = Get-Content -LiteralPath $FollowerCacheJson -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($property in $cacheObject.PSObject.Properties) {
        $cache[$property.Name] = $property.Value
    }
}

function Save-FollowerCache {
    $ordered = [ordered]@{}
    foreach ($key in ($cache.Keys | Sort-Object)) {
        $ordered[$key] = $cache[$key]
    }
    $ordered | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $FollowerCacheJson -Encoding UTF8
}

function Get-FollowerCount {
    param([Parameter(Mandatory = $true)][string]$Uid)

    if ($cache.ContainsKey($Uid) -and [string]$cache[$Uid].核验日 -eq $today -and $null -ne $cache[$Uid].粉丝数) {
        return $cache[$Uid]
    }
    if ($SkipFollowerFetch) {
        return $null
    }

    Start-Sleep -Milliseconds $MinDelayMs
    try {
        $uri = "https://api.bilibili.com/x/relation/stat?vmid=$Uid"
        $response = Invoke-RestMethod -Uri $uri -Headers $requestHeaders -TimeoutSec 30
        if ([int]$response.code -ne 0) {
            throw "Bilibili API code $($response.code): $($response.message)"
        }
        $entry = [pscustomobject]@{
            粉丝数 = [long]$response.data.follower
            核验日 = $today
            错误   = ''
        }
        $cache[$Uid] = $entry
        Save-FollowerCache
        return $entry
    }
    catch {
        $message = $_.Exception.Message
        $cache[$Uid] = [pscustomobject]@{
            粉丝数 = $null
            核验日 = $today
            错误   = $message
        }
        Save-FollowerCache
        if ($message -match '\b412\b') {
            throw
        }
        return $cache[$Uid]
    }
}

$areaToCity = @{
    '无锡市' = '无锡市'; '江阴市' = '无锡市'; '宜兴市' = '无锡市'
    '苏州市' = '苏州市'; '常熟市' = '苏州市'; '张家港市' = '苏州市'; '昆山市' = '苏州市'; '太仓市' = '苏州市'
    '常州市' = '常州市'; '溧阳市' = '常州市'
    '镇江市' = '镇江市'; '丹阳市' = '镇江市'; '扬中市' = '镇江市'; '句容市' = '镇江市'
    '溧水区' = '南京市'
}

function Resolve-RowArea {
    param([Parameter(Mandatory = $true)]$Row)
    $title = [string]$Row.title
    $description = [string]$Row.description
    $tags = [string]$Row.tag
    $areas = @(
        @($Row.areas) + @([string]$Row.area) |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) -and $areaToCity.ContainsKey([string]$_) } |
            Sort-Object @{ Expression = { (([string]$_) -replace '市$', '').Length }; Descending = $true }, @{ Expression = { [string]$_ } } -Unique
    )
    foreach ($text in @($title, $description, $tags)) {
        foreach ($area in $areas) {
            $token = ([string]$area) -replace '市$', ''
            if ($text.Contains($token)) {
                if ($text -eq $title) {
                    return [pscustomobject]@{ Area = [string]$area; City = [string]$areaToCity[[string]$area] }
                }
                break
            }
        }
        if ($text -eq $title) {
            $outsideCity = '北京|上海|南京|杭州|宁波|扬州|泰州|南通|嘉兴|湖州|绍兴|盐城|徐州|淮安|宿迁|连云港|合肥|武汉|成都|重庆|广州|深圳'
            if ($title -match $outsideCity) {
                return $null
            }
        }
        foreach ($area in $areas) {
            $token = ([string]$area) -replace '市$', ''
            if ($text.Contains($token)) {
                return [pscustomobject]@{ Area = [string]$area; City = [string]$areaToCity[[string]$area] }
            }
        }
    }
    return $null
}

function Test-AreaMatch {
    param([Parameter(Mandatory = $true)]$Row)
    return $null -ne (Resolve-RowArea -Row $Row)
}

function Test-FoodMatch {
    param([Parameter(Mandatory = $true)]$Row)
    $content = "{0} {1}" -f [string]$Row.title, [string]$Row.description
    $tags = [string]$Row.tag
    $nonVisitNews = '打人|殴打|冲突|纠纷|事件起因|警方通报|执法|维权|曝光商家|食品安全通报'
    if ($content -match $nonVisitNews -and $content -notmatch '探店|到店|打卡|品尝|试吃|测评|逛吃') {
        return $false
    }
    $strongFood = '美食|小吃|饭店|餐厅|餐馆|饭馆|菜馆|面馆|早餐|早茶|夜宵|烧烤|烤肉|自助|火锅|本帮菜|淮扬菜|小笼|包子|馄饨|食堂|餐饮|饭局|酒楼|饭庄|料理|放题|吃播'
    $dishOrEating = '吃|菜|饭|面|粉|汤|饼|糕|馒头|鸭|鸡|鱼|肉|虾|蟹|牛排|汉堡|咖啡|奶茶|甜品|点心|串|锅|粥'
    return $content -match $strongFood -or
        ($content -match '探店' -and $content -match $dishOrEating) -or
        ($tags -match $strongFood -and $content -match $dishOrEating)
}

$eligibleUids = @(
    $uniqueRows |
        Where-Object { (Test-AreaMatch $_) -and (Test-FoodMatch $_) -and [long]$_.comments -gt $MinComments } |
        ForEach-Object { [string]$_.mid } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
)

$checked = 0
foreach ($uid in $eligibleUids) {
    $null = Get-FollowerCount -Uid $uid
    $checked++
    if ($checked % 25 -eq 0 -or $checked -eq $eligibleUids.Count) {
        [pscustomobject]@{
            已核验 = $checked
            总数   = $eligibleUids.Count
            UID    = $uid
        } | ConvertTo-Json -Compress
    }
}

$oldStatus = @{}
if (Test-Path -LiteralPath $SourceCsv) {
    foreach ($row in @(Import-Csv -LiteralPath $SourceCsv -Encoding UTF8)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$row.BV号)) {
            $oldStatus[[string]$row.BV号] = $row
        }
    }
}

$batchStatus = @{}
if (Test-Path -LiteralPath $CommentStatusCsv) {
    foreach ($row in @(Import-Csv -LiteralPath $CommentStatusCsv -Encoding UTF8)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$row.BV号)) {
            $batchStatus[[string]$row.BV号] = $row
        }
    }
}

$candidateRows = foreach ($row in $uniqueRows) {
    $uid = [string]$row.mid
    $follower = if ($cache.ContainsKey($uid)) { $cache[$uid] } else { $null }
    $followerCount = if ($null -ne $follower -and $null -ne $follower.粉丝数) { [long]$follower.粉丝数 } else { $null }
    $commentPass = [long]$row.comments -gt $MinComments
    $followerPass = $null -ne $followerCount -and $followerCount -gt $MinFollowers
    $resolvedArea = Resolve-RowArea -Row $row
    $areaPass = $null -ne $resolvedArea
    $foodPass = Test-FoodMatch $row
    $conclusion = if (-not $areaPass) {
        '不通过：缺少目标地区线索'
    }
    elseif (-not $foodPass) {
        '不通过：缺少餐饮探店语义'
    }
    elseif (-not $commentPass) {
        "不通过：评论不超过$MinComments"
    }
    elseif ($null -eq $followerCount) {
        '待核验：粉丝数未取得'
    }
    elseif (-not $followerPass) {
        "不通过：粉丝不超过$MinFollowers"
    }
    else {
        "通过：粉丝>$MinFollowers 且评论>$MinComments"
    }

    $published = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$row.pubdate)) {
        $published = [DateTimeOffset]::FromUnixTimeSeconds([long]$row.pubdate).ToOffset([TimeSpan]::FromHours(8)).ToString('yyyy-MM-dd')
    }

    [pscustomobject]@{
        城市范围       = if ($areaPass) { [string]$resolvedArea.City } else { [string]$row.city }
        区县或县级市   = if ($areaPass) { [string]$resolvedArea.Area } else { [string]$row.area }
        检索日期       = $today
        检索词         = (@($row.queries) -join '；')
        检索页码       = (@($row.pages | Sort-Object -Unique) -join '；')
        视频标题       = [string]$row.title
        BV号           = [string]$row.bvid
        视频链接       = "https://www.bilibili.com/video/$($row.bvid)/"
        发布日期       = $published
        播放量         = [long]$row.play
        评论数         = [long]$row.comments
        弹幕数         = [long]$row.danmaku
        创作者         = [string]$row.author
        创作者UID      = $uid
        粉丝数         = if ($null -eq $followerCount) { '' } else { $followerCount }
        粉丝核验日     = if ($null -eq $follower) { '' } else { [string]$follower.核验日 }
        地区匹配       = if ($areaPass) { '通过' } else { '不通过' }
        餐饮语义       = if ($foodPass) { '通过' } else { '不通过' }
        评论门槛       = if ($commentPass) { "通过：>$MinComments" } else { "不通过：<=$MinComments" }
        粉丝门槛       = if ($null -eq $followerCount) { '待核验' } elseif ($followerPass) { "通过：>$MinFollowers" } else { "不通过：<=$MinFollowers" }
        门槛结论       = $conclusion
        发现来源       = 'B站内置浏览器搜索结果逐页采集'
    }
}

$candidateRows = @($candidateRows | Sort-Object 城市范围, 区县或县级市, @{ Expression = { [long]$_.评论数 }; Descending = $true }, BV号)
$candidateRows | Export-Csv -LiteralPath $CandidateCsv -NoTypeInformation -Encoding UTF8

$sourceRows = foreach ($row in $candidateRows | Where-Object { $_.门槛结论 -like '通过：*' }) {
    $old = if ($oldStatus.ContainsKey([string]$row.BV号)) { $oldStatus[[string]$row.BV号] } else { $null }
    $rawPath = Join-Path $CommentDir "$($row.BV号).json"
    $batch = if ($batchStatus.ContainsKey([string]$row.BV号)) { $batchStatus[[string]$row.BV号] } else { $null }
    $commentStatus = if (Test-Path -LiteralPath $rawPath) {
        try {
            $rawComment = Get-Content -LiteralPath $rawPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([bool]$rawComment.完整) {
                "完整（$($rawComment.实际抓取数)/$($rawComment.页面评论总数)，含回复）"
            }
            else {
                "未完整（$($rawComment.实际抓取数)/$($rawComment.页面评论总数)，含回复）"
            }
        }
        catch {
            '已有原始JSON，但解析失败，待复核'
        }
    }
    elseif ($null -ne $batch -and [string]$batch.采集状态 -eq '失败') {
        if ([string]$batch.错误 -match '412') { '失败：HTTP 412，已停止批次' } else { "失败：$([string]$batch.错误)" }
    }
    elseif ($null -ne $old -and -not [string]::IsNullOrWhiteSpace([string]$old.评论采集状态)) {
        [string]$old.评论采集状态
    }
    else {
        '待重新采集'
    }
    [pscustomobject]@{
        城市           = $row.城市范围
        区县或县级市   = $row.区县或县级市
        检索日期       = $row.检索日期
        视频标题       = $row.视频标题
        BV号           = $row.BV号
        视频链接       = $row.视频链接
        发布日         = $row.发布日期
        播放量         = $row.播放量
        评论数         = $row.评论数
        创作者         = $row.创作者
        创作者UID      = $row.创作者UID
        粉丝数         = $row.粉丝数
        粉丝核验日     = $row.粉丝核验日
        来源           = 'B站内置浏览器搜索页结构化数据+公开粉丝接口'
        门槛结论       = $row.门槛结论
        评论采集状态   = $commentStatus
        备注           = if ($null -eq $old) { '地区和餐饮语义由标题、标签或简介命中；评论数为搜索页结构化字段 review' } else { [string]$old.备注 }
    }
}

$sourceRows = @($sourceRows | Sort-Object 城市, 区县或县级市, @{ Expression = { [long]$_.评论数 }; Descending = $true }, BV号)
$sourceRows | Export-Csv -LiteralPath $SourceCsv -NoTypeInformation -Encoding UTF8

$runRows = foreach ($run in $runs) {
    [pscustomobject]@{
        城市         = [string]$run.city
        检索单元     = [string]$run.area
        检索词       = [string]$run.term
        页面公布末页 = [int]$run.lastPageAdvertised
        实际末页     = [int]$run.lastPageFetched
        原始卡片数   = [int]$run.rawRows
        停止原因     = [string]$run.stopReason
        失败或空页数 = @($run.failures).Count
        检索日期     = $today
    }
}
$runRows | Export-Csv -LiteralPath $DiscoveryRunCsv -NoTypeInformation -Encoding UTF8

[pscustomobject]@{
    原始唯一BV数 = $uniqueRows.Count
    待粉丝核验UID数 = $eligibleUids.Count
    候选行数 = $candidateRows.Count
    正式来源行数 = $sourceRows.Count
    候选表 = $CandidateCsv
    来源表 = $SourceCsv
    批次表 = $DiscoveryRunCsv
    粉丝缓存 = $FollowerCacheJson
} | ConvertTo-Json -Compress
