[CmdletBinding()]
param(
    [string]$RawDirectory = (Join-Path ([IO.Path]::GetTempPath()) 'healthy_life_bili_comments'),
    [string]$AliasCsv = '',
    [string]$EvidenceCsv = '',
    [string]$VerificationCsv = '',
    [string]$SourceCsv = '',
    [string]$CandidateOutput = '',
    [string]$EntityOutput = ''
)

$ErrorActionPreference = 'Stop'
$dataDirectory = Join-Path $PSScriptRoot '..\饭店数据'
if (-not $AliasCsv) { $AliasCsv = Join-Path $dataDirectory '饭店别名词典.csv' }
if (-not $EvidenceCsv) { $EvidenceCsv = Join-Path $dataDirectory '评论推荐.csv' }
if (-not $VerificationCsv) { $VerificationCsv = Join-Path $dataDirectory '门店核验.csv' }
if (-not $SourceCsv) { $SourceCsv = Join-Path $dataDirectory '视频来源.csv' }
if (-not $CandidateOutput) { $CandidateOutput = Join-Path $dataDirectory '饭店提及候选.csv' }
if (-not $EntityOutput) { $EntityOutput = Join-Path $dataDirectory '饭店实体精筛.csv' }

function New-StringSet {
    New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
}

function Limit-Text([string]$Text, [int]$Length = 120) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $oneLine = ($Text -replace '\s+', ' ').Trim()
    if ($oneLine.Length -le $Length) { return $oneLine }
    return $oneLine.Substring(0, $Length) + '…'
}

function Clean-Candidate([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $name = $Value -replace '\[[^\]]{1,20}\]', ''
    $name = $name -replace '^[#@＠]+', ''
    $name = $name.Trim(' ', "`t", '，', ',', '。', '.', '！', '!', '？', '?', '：', ':', '；', ';', '、', '“', '”', '「', '」', '《', '》', '（', '）', '(', ')', '-', '—', '~', '～')
    $name = $name -replace '^(?:我觉得|个人觉得|感觉|其实|真的|直接|本地人|无锡人|无锡的|无锡|还是|就是|这个|那个|这家|那家|一家|推荐去|推荐吃|推荐|建议去|建议吃|建议|可以去|去吃|去尝尝|试试|首选|常去|必吃|安利|强推|店叫|叫做|叫)+', ''
    $name = $name -replace '^(?:吃|去)+', ''
    $name = ($name -split '(?:比这家|比这个|比视频里|更好吃|更好|更强|也不错|不错|好吃|值得|便宜|太贵|很贵|难吃|一般|不行|可以|推荐|老板|店里|开了|关了|倒闭|搬走)')[0]
    $name = $name -replace '(?:的店|这家店|那家店|老店面馆|老店)$', { param($m) if ($m.Value -eq '老店面馆') { '面馆' } else { '' } }
    $name = $name.Trim()

    if ($name.Length -lt 2 -or $name.Length -gt 18) { return $null }
    if ($name -notmatch '[A-Za-z一-龥]') { return $null }
    if ($name -match '^(?:饭店|餐厅|面馆|馄饨|小笼|包子|烧烤|火锅|小吃|早餐|夜宵|本地人|无锡人|老板|博主|视频|评论|这里|那里|这边|那边|一家店|老字号|学校食堂|单位食堂|食堂|我家|你家|他家|自己家|家里|附近|市区|南长街|惠山古镇|不知道|没吃过|都可以|随便|下次|小时候|以前|现在|确实|感觉|个人|朋友|同事|外地人)$') { return $null }
    if ($name -match '^(?:这个|那个|这种|那种|什么|怎么|为什么|有没有|哪里|哪个|一个|很多|好多|真正|所谓|普通|正宗|当地|本地|无锡|江阴|宜兴)') { return $null }
    if ($name -match '(?:饭店不干了|馆子|店名|饭店里|餐厅里|面馆里|食堂里|视频里|评论区|博主|老板说|服务员|我家楼|学校|医院|公司|小区)$') { return $null }
    return $name
}

function New-Aggregate([string]$Name) {
    [pscustomobject]@{
        Name = $Name
        CommentKeys = New-StringSet
        Videos = New-StringSet
        Cities = New-StringSet
        Areas = New-StringSet
        Methods = New-StringSet
        PositiveKeys = New-StringSet
        NegativeKeys = New-StringSet
        BestScore = -1
        EvidenceBv = ''
        EvidenceId = ''
        EvidenceText = ''
        MaxLikes = 0
    }
}

function Add-Hit($Aggregate, $Comment, [string]$Method, [bool]$Positive, [bool]$Negative) {
    $key = [string]$Comment.Key
    [void]$Aggregate.CommentKeys.Add($key)
    [void]$Aggregate.Videos.Add([string]$Comment.Bv)
    $source = $sourceByBvid[[string]$Comment.Bv]
    if ($null -ne $source) {
        if (-not [string]::IsNullOrWhiteSpace([string]$source.城市)) { [void]$Aggregate.Cities.Add([string]$source.城市) }
        if (-not [string]::IsNullOrWhiteSpace([string]$source.区县或县级市)) { [void]$Aggregate.Areas.Add([string]$source.区县或县级市) }
    }
    [void]$Aggregate.Methods.Add($Method)
    if ($Positive) { [void]$Aggregate.PositiveKeys.Add($key) }
    if ($Negative) { [void]$Aggregate.NegativeKeys.Add($key) }
    $likes = [int]$Comment.Likes
    if ($likes -gt $Aggregate.MaxLikes) { $Aggregate.MaxLikes = $likes }
    $score = $likes + $(if ($Positive) { 100 } else { 0 }) + $(if ($Method -eq '词典命中') { 50 } elseif ($Method -eq '比较／推荐句式') { 30 } else { 10 })
    if ($score -gt $Aggregate.BestScore) {
        $Aggregate.BestScore = $score
        $Aggregate.EvidenceBv = [string]$Comment.Bv
        $Aggregate.EvidenceId = [string]$Comment.Id
        $Aggregate.EvidenceText = Limit-Text ([string]$Comment.Text)
    }
}

if (-not (Test-Path -LiteralPath $RawDirectory -PathType Container)) {
    throw "原始评论目录不存在：$RawDirectory"
}
foreach ($required in @($AliasCsv, $EvidenceCsv, $VerificationCsv, $SourceCsv)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) { throw "缺少输入文件：$required" }
}

$dictionary = @(Import-Csv -LiteralPath $AliasCsv -Encoding UTF8)
$verification = @(Import-Csv -LiteralPath $VerificationCsv -Encoding UTF8)
$sourceByBvid = @{}
foreach ($sourceRow in @(Import-Csv -LiteralPath $SourceCsv -Encoding UTF8)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$sourceRow.BV号)) {
        $sourceByBvid[[string]$sourceRow.BV号] = $sourceRow
    }
}
$commentsByKey = @{}
$rawFileCount = 0

Get-ChildItem -LiteralPath $RawDirectory -Filter '*.json' -File | Sort-Object Name | ForEach-Object {
    $rawFileCount++
    $payload = Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($row in @($payload.评论)) {
        $key = "{0}|{1}" -f [string]$row.BV号, [string]$row.评论ID
        if (-not $commentsByKey.ContainsKey($key)) {
            $commentsByKey[$key] = [pscustomobject]@{
                Key = $key
                Bv = [string]$row.BV号
                Id = [string]$row.评论ID
                ParentId = [string]$row.父评论ID
                Published = [string]$row.发布时间
                Likes = [int]$row.点赞数
                Text = [string]$row.评论文本
                SourceType = '原始评论'
            }
        }
    }
}

$selectedEvidenceAdded = 0
foreach ($row in @(Import-Csv -LiteralPath $EvidenceCsv -Encoding UTF8)) {
    $key = "{0}|{1}" -f [string]$row.BV号, [string]$row.评论ID
    if (-not $commentsByKey.ContainsKey($key)) {
        $commentsByKey[$key] = [pscustomobject]@{
            Key = $key
            Bv = [string]$row.BV号
            Id = [string]$row.评论ID
            ParentId = [string]$row.父评论ID
            Published = [string]$row.发布时间
            Likes = $(if ([string]::IsNullOrWhiteSpace($row.点赞数)) { 0 } else { [int]$row.点赞数 })
            Text = [string]$row.评论证据摘要
            SourceType = '历史精选证据摘要'
        }
        $selectedEvidenceAdded++
    }
}

$aliasRows = @()
$canonicalAggregates = @{}
foreach ($entry in $dictionary) {
    $canonical = [string]$entry.标准店名
    $canonicalAggregates[$canonical] = New-Aggregate $canonical
    foreach ($alias in @(([string]$entry.别名列表) -split '\|' | Where-Object { $_ })) {
        $aliasRows += [pscustomobject]@{ Canonical = $canonical; Alias = [string]$alias; Entry = $entry }
    }
}
$aliasRows = @($aliasRows | Sort-Object @{ Expression = { $_.Alias.Length }; Descending = $true })

$candidateAggregates = @{}
$positivePattern = '(?:推荐|建议|值得|好吃|不错|可以去|去吃|去尝尝|试试|首选|常去|必吃|安利|强推|更好|更强|吊打|完爆|不如)'
$negativePattern = '(?:不推荐|不好吃|难吃|踩雷|避雷|失望|太贵|很贵|坑人|不行|一般|倒闭|关门|关了)'
$cuePattern = '(?:还不如|不如|推荐去|推荐吃|推荐|建议去|建议吃|建议|可以去|去吃|去尝尝|试试|首选|常去|必吃|安利|强推|店叫|叫做|叫)\s*(?<name>[A-Za-z0-9一-龥·&]{2,18})'
$beforeCuePattern = '(?<name>[A-Za-z0-9一-龥·&]{2,18})(?:比这家|比这个|比视频里|更好吃|更好|更强|值得去|值得吃)'
$suffixPattern = '[A-Za-z0-9一-龥·&]{2,18}(?:鸡鸭鹅大酒店|鸡鸭鹅大饭店|大酒店|大饭店|鸡汤面馆|馄饨店|酸辣汤|铁板烧|卤肉饭|酸菜鱼|猪脚饭|海棠糕|面馆|饭店|餐厅|食府|酒楼|饭庄|小馆|烤肉|烧烤|火锅|馄饨|小笼|包子|糕团|点心|炸鸡|拉面|生煎|小吃|酒店|菜馆)'

foreach ($comment in @($commentsByKey.Values)) {
    $text = ([string]$comment.Text -replace '\[[^\]]{1,20}\]', '') -replace '\s+', ' '
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    $isPositive = $text -match $positivePattern
    $isNegative = $text -match $negativePattern
    $commentCandidates = @{}

    foreach ($aliasRow in $aliasRows) {
        if ($text.IndexOf($aliasRow.Alias, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            Add-Hit $canonicalAggregates[$aliasRow.Canonical] $comment '词典命中' $isPositive $isNegative
            $rawKey = $aliasRow.Alias.ToLowerInvariant()
            if (-not $commentCandidates.ContainsKey($rawKey)) {
                $commentCandidates[$rawKey] = [pscustomobject]@{ Name = $aliasRow.Alias; Methods = New-StringSet }
            }
            [void]$commentCandidates[$rawKey].Methods.Add('词典命中')
        }
    }

    foreach ($match in [regex]::Matches($text, $cuePattern, 'IgnoreCase')) {
        $name = Clean-Candidate $match.Groups['name'].Value
        if ($name) {
            $key = $name.ToLowerInvariant()
            if (-not $commentCandidates.ContainsKey($key)) { $commentCandidates[$key] = [pscustomobject]@{ Name = $name; Methods = New-StringSet } }
            [void]$commentCandidates[$key].Methods.Add('比较／推荐句式')
        }
    }
    foreach ($match in [regex]::Matches($text, $beforeCuePattern, 'IgnoreCase')) {
        $name = Clean-Candidate $match.Groups['name'].Value
        if ($name) {
            $key = $name.ToLowerInvariant()
            if (-not $commentCandidates.ContainsKey($key)) { $commentCandidates[$key] = [pscustomobject]@{ Name = $name; Methods = New-StringSet } }
            [void]$commentCandidates[$key].Methods.Add('比较／推荐句式')
        }
    }
    foreach ($segment in @($text -split '[\s，,。.!！?？；;：:、/／|（）()【】\[\]“”"''<>《》]+')) {
        if ($segment.Length -lt 2 -or $segment.Length -gt 40) { continue }
        foreach ($match in [regex]::Matches($segment, $suffixPattern, 'IgnoreCase')) {
            $name = Clean-Candidate $match.Value
            if ($name) {
                $key = $name.ToLowerInvariant()
                if (-not $commentCandidates.ContainsKey($key)) { $commentCandidates[$key] = [pscustomobject]@{ Name = $name; Methods = New-StringSet } }
                [void]$commentCandidates[$key].Methods.Add('门店后缀')
            }
        }
    }

    foreach ($candidate in @($commentCandidates.Values)) {
        $key = $candidate.Name.ToLowerInvariant()
        if (-not $candidateAggregates.ContainsKey($key)) { $candidateAggregates[$key] = New-Aggregate $candidate.Name }
        foreach ($method in $candidate.Methods) { Add-Hit $candidateAggregates[$key] $comment $method $isPositive $isNegative }
    }
}

$candidateRows = foreach ($aggregate in $candidateAggregates.Values) {
    $count = $aggregate.CommentKeys.Count
    $isDictionary = $aggregate.Methods.Contains('词典命中')
    $hasCue = $aggregate.Methods.Contains('比较／推荐句式')
    $keep = $isDictionary -or $count -ge 2 -or ($hasCue -and $aggregate.Name.Length -le 12)
    if (-not $keep) { continue }
    $level = if ($isDictionary) { 'A-词典命中' } elseif ($count -ge 2 -and $hasCue) { 'B-多条且有推荐语境' } elseif ($count -ge 2) { 'C-多条提及' } else { 'D-单条宽松召回' }
    [pscustomobject]@{
        候选写法 = $aggregate.Name
        提及评论数 = $count
        涉及视频数 = $aggregate.Videos.Count
        涉及城市 = (($aggregate.Cities | Sort-Object) -join '／')
        涉及区县或县级市 = (($aggregate.Areas | Sort-Object) -join '／')
        推荐或比较语境数 = $aggregate.PositiveKeys.Count
        负面语境数 = $aggregate.NegativeKeys.Count
        最高点赞数 = $aggregate.MaxLikes
        召回方式 = (($aggregate.Methods | Sort-Object) -join '／')
        召回级别 = $level
        证据BV号 = $aggregate.EvidenceBv
        证据评论ID = $aggregate.EvidenceId
        证据摘录 = $aggregate.EvidenceText
        处理建议 = $(if ($isDictionary) { '进入实体精筛表' } elseif ($hasCue -or $count -ge 3) { '优先平台检索' } else { '人工复核后再检索' })
    }
}
$candidateRows = @($candidateRows | Sort-Object @{ Expression = '提及评论数'; Descending = $true }, @{ Expression = '推荐或比较语境数'; Descending = $true }, 候选写法)

$entityRows = foreach ($entry in $dictionary) {
    $canonical = [string]$entry.标准店名
    $aggregate = $canonicalAggregates[$canonical]
    if ($aggregate.CommentKeys.Count -eq 0) { continue }
    $aliases = @(([string]$entry.别名列表) -split '\|' | Where-Object { $_ })
    $matchedVerification = @($verification | Where-Object {
        $verifiedName = [string]$_.标准店名
        $verifiedName -eq $canonical -or $aliases -contains $verifiedName
    })
    $mapLinks = @($matchedVerification | ForEach-Object { $_.高德链接 } | Where-Object { $_ } | Sort-Object -Unique)
    $dianpingLinks = @($matchedVerification | ForEach-Object { $_.点评链接 } | Where-Object { $_ } | Sort-Object -Unique)
    $platformState = if ($mapLinks.Count -and $dianpingLinks.Count) { '双平台已有记录' } elseif ($mapLinks.Count) { '仅地图已有记录' } elseif ($dianpingLinks.Count) { '仅点评已有记录' } else { '待平台核验' }
    $stage = if ([string]$entry.精筛结论 -eq '待歧义消解') { 'D-歧义待解' } elseif ($platformState -eq '双平台已有记录') { 'P2-双平台匹配' } elseif ($platformState -match '^仅') { 'P1-单平台匹配' } else { 'N-待平台核验' }
    [pscustomobject]@{
        标准店名 = $canonical
        出现写法 = ([string]$entry.别名列表 -replace '\|', '／')
        评论提及数 = $aggregate.CommentKeys.Count
        涉及视频数 = $aggregate.Videos.Count
        涉及城市 = (($aggregate.Cities | Sort-Object) -join '／')
        涉及区县或县级市 = (($aggregate.Areas | Sort-Object) -join '／')
        推荐或比较语境数 = $aggregate.PositiveKeys.Count
        负面语境数 = $aggregate.NegativeKeys.Count
        精筛结论 = [string]$entry.精筛结论
        类别 = [string]$entry.类别
        平台匹配状态 = $platformState
        高德链接 = ($mapLinks -join ' | ')
        点评链接 = ($dianpingLinks -join ' | ')
        精筛阶段 = $stage
        证据BV号 = $aggregate.EvidenceBv
        证据评论ID = $aggregate.EvidenceId
        证据摘录 = $aggregate.EvidenceText
        备注 = [string]$entry.备注
    }
}
$entityRows = @($entityRows | Sort-Object @{ Expression = '评论提及数'; Descending = $true }, 标准店名)

$candidateParent = Split-Path -Parent $CandidateOutput
$entityParent = Split-Path -Parent $EntityOutput
New-Item -ItemType Directory -Force -Path $candidateParent, $entityParent | Out-Null
$candidateRows | Export-Csv -LiteralPath $CandidateOutput -NoTypeInformation -Encoding UTF8
$entityRows | Export-Csv -LiteralPath $EntityOutput -NoTypeInformation -Encoding UTF8

[pscustomobject]@{
    原始JSON文件数 = $rawFileCount
    原始评论数 = @($commentsByKey.Values | Where-Object { $_.SourceType -eq '原始评论' }).Count
    新增历史精选证据数 = $selectedEvidenceAdded
    实际可扫描记录数 = $commentsByKey.Count
    宽松候选写法数 = $candidateRows.Count
    精筛实体数 = $entityRows.Count
    候选输出 = (Resolve-Path -LiteralPath $CandidateOutput).Path
    精筛输出 = (Resolve-Path -LiteralPath $EntityOutput).Path
} | Format-List
