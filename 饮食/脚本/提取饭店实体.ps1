[CmdletBinding()]
param(
    [string]$RawDirectory = (Join-Path ([IO.Path]::GetTempPath()) 'healthy_life_bili_comments'),
    [string]$AliasCsv = '',
    [string]$EvidenceCsv = '',
    [string]$VerificationCsv = '',
    [string]$SourceCsv = '',
    [string]$SemanticRuleCsv = '',
    [string]$CandidateAuditCsv = '',
    [string]$CandidateOutput = '',
    [string]$EntityOutput = '',
    [string]$SemanticOutput = '',
    [string]$RecommendedOutput = '',
    [switch]$SemanticSelfTest
)

$ErrorActionPreference = 'Stop'
$dataDirectory = Join-Path $PSScriptRoot '..\饭店数据'
if (-not $AliasCsv) { $AliasCsv = Join-Path $dataDirectory '饭店别名词典.csv' }
if (-not $EvidenceCsv) { $EvidenceCsv = Join-Path $dataDirectory '评论推荐.csv' }
if (-not $VerificationCsv) { $VerificationCsv = Join-Path $dataDirectory '门店核验.csv' }
if (-not $SourceCsv) { $SourceCsv = Join-Path $dataDirectory '视频来源.csv' }
if (-not $SemanticRuleCsv) { $SemanticRuleCsv = Join-Path $dataDirectory '饭店语义规则.csv' }
if (-not $CandidateAuditCsv) { $CandidateAuditCsv = Join-Path $dataDirectory '饭店候选人工审计.csv' }
if (-not $CandidateOutput) { $CandidateOutput = Join-Path $dataDirectory '饭店提及候选.csv' }
if (-not $EntityOutput) { $EntityOutput = Join-Path $dataDirectory '饭店实体精筛.csv' }
if (-not $SemanticOutput) { $SemanticOutput = Join-Path $dataDirectory '饭店推荐语义标注.csv' }
if (-not $RecommendedOutput) { $RecommendedOutput = Join-Path $dataDirectory '饭店有效推荐评论.csv' }

function New-StringSet {
    New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
}

function Remove-UserReferences([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $sanitized = $Text -replace '回复\s+@[^\s:：，,。！？!?；;]+\s*[:：]?', '回复：'
    $sanitized = $sanitized -replace '@[^\s:：，,。！？!?；;]+', '@用户'
    return ($sanitized -replace '\s+', ' ').Trim()
}

function Limit-Text([string]$Text, [int]$Length = 120) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $oneLine = ((Remove-UserReferences $Text) -replace '\s+', ' ').Trim()
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
    $name = $name -replace '(?:的店|这家店|那家店|老店面馆|老店|去吃|吃)$', { param($m) if ($m.Value -eq '老店面馆') { '面馆' } else { '' } }
    $name = $name.Trim()

    if ($name.Length -lt 2 -or $name.Length -gt 18) { return $null }
    if ($name -notmatch '[A-Za-z一-龥]') { return $null }
    if ($name -match '^(?:饭店|餐厅|面馆|馄饨|小笼|包子|烧烤|火锅|小吃|早餐|夜宵|本地人|无锡人|老板|博主|视频|评论|这里|那里|这边|那边|一家店|老字号|学校食堂|单位食堂|食堂|我家|你家|他家|自己家|家里|附近|市区|南长街|惠山古镇|不知道|没吃过|都可以|随便|下次|小时候|以前|现在|确实|感觉|个人|朋友|同事|外地人)$') { return $null }
    if ($name -match '^(?:这个|那个|这种|那种|什么|怎么|为什么|有没有|哪里|哪个|一个|很多|好多|真正|所谓|普通|正宗|当地|本地|无锡|江阴|宜兴)') { return $null }
    if ($name -match '(?:饭店不干了|馆子|店名|饭店里|餐厅里|面馆里|食堂里|视频里|评论区|博主|老板说|服务员|我家楼|学校|医院|公司|小区)$') { return $null }
    if ($name -match '^京东') { return $null }
    return $name
}

function Test-ContainsAny([string]$Text, [string[]]$Terms) {
    foreach ($term in $Terms) {
        if ($Text.IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

$semanticRecommendCues = @('强烈推荐', '最推荐', '强推', '推荐去', '推荐吃', '我推荐', '推荐', '建议去', '建议吃', '值得去', '值得吃', '可以去', '可以吃', '可以试试', '安利', '必吃', '首选')
$semanticPositiveCues = @('很好吃', '真好吃', '好吃', '很不错', '也不错', '确实不错', '确实也不错', '相当可以', '味道很好', '味道不错', '质量在线', '性价比高', '价格实惠', '很实惠', '很喜欢', '喜欢吃', '值得一试', '值得尝尝', '稳定', '正宗', '地道')
$semanticNegativeCues = @('不推荐', '不好吃', '难吃', '踩雷', '避雷', '劝退', '失望', '太贵', '很贵', '坑人', '不行', '一般', '倒闭', '关门', '关了', '再也不去', '不会再去')
$semanticComparisonCues = @('还不如', '不如', '比不上', '不及', '还不及')
$semanticIronyCues = @('[疑惑]', '[无语]', '笑死', '呵呵', '就这', '打广告', '装死', '洗地', '吹捧', '阴阳', '反向推荐')
$semanticBehaviorCues = @('经常去', '常去', '一直去', '每次去', '聚餐都在', '现在都在', '反复去', '回购')
$semanticEatingCues = @('吃', '点', '尝', '打卡', '聚餐')

function Get-SemanticAssessment([string]$Text, [string[]]$Names) {
    $result = [pscustomobject]@{
        Recommended = $false
        Comparison = $false
        Negative = $false
        MentionOnly = $false
        Reasons = New-StringSet
    }
    if ([string]::IsNullOrWhiteSpace($Text)) {
        $result.MentionOnly = $true
        [void]$result.Reasons.Add('只有店名提及')
        return $result
    }

    $sentences = @($Text.Split([char[]]"。！？!?；;`r`n", [StringSplitOptions]::RemoveEmptyEntries))
    foreach ($sentence in $sentences) {
        $listScope = $false
        $sentenceApprovesList = (Test-ContainsAny $sentence @('这几个都可以', '这些都可以', '这几家都可以', '都不错', '都挺好', '都值得试试')) -and -not (Test-ContainsAny $sentence @('都不推荐', '都不好吃', '都不行'))
        $colonIndex = $sentence.IndexOfAny([char[]]'：:')
        if ($colonIndex -ge 0) {
            $prefix = $sentence.Substring(0, $colonIndex)
            $listScope = (Test-ContainsAny $prefix $semanticRecommendCues) -and -not (Test-ContainsAny $prefix @('不推荐', '别去', '不要去'))
        }
        foreach ($clause in @($sentence.Split([char[]]'，,、', [StringSplitOptions]::RemoveEmptyEntries))) {
            $hasComparison = Test-ContainsAny $clause $semanticComparisonCues
            $hasIrony = Test-ContainsAny $clause $semanticIronyCues
            $hasNegative = Test-ContainsAny $clause $semanticNegativeCues
            if ($clause.Contains('不难吃')) { $hasNegative = $false }
            $hasDirectRecommendation = Test-ContainsAny $clause $semanticRecommendCues
            $hasPositiveExperience = Test-ContainsAny $clause $semanticPositiveCues
            $hasRepeatedBehavior = (Test-ContainsAny $clause $semanticBehaviorCues) -and (Test-ContainsAny $clause $semanticEatingCues)
            if ($hasNegative -or $hasIrony -or ($hasComparison -and -not $hasDirectRecommendation)) {
                $listScope = $false
            }
            elseif ($hasDirectRecommendation) {
                $listScope = $true
            }

            $matched = $false
            foreach ($name in $Names) {
                if (-not [string]::IsNullOrWhiteSpace($name) -and $clause.IndexOf($name, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $matched = $true
                    break
                }
            }
            if (-not $matched) { continue }

            if ($hasComparison -and -not $hasDirectRecommendation) {
                $result.Comparison = $true
                [void]$result.Reasons.Add('戏谑或比较提及，不等同推荐')
            }
            if ($hasNegative -or $hasIrony) {
                $result.Negative = $true
                [void]$result.Reasons.Add($(if ($hasIrony) { '含反讽或戏谑信号' } else { '含明确负向体验' }))
            }
            $genuineRecommendation = ($hasDirectRecommendation -or $hasPositiveExperience -or $hasRepeatedBehavior -or $listScope -or $sentenceApprovesList) -and -not $hasNegative -and -not $hasIrony
            if ($hasComparison -and -not $hasDirectRecommendation) { $genuineRecommendation = $false }
            if ($genuineRecommendation) {
                $result.Recommended = $true
                $reason = if ($hasDirectRecommendation) { '明确到店或菜品推荐' } elseif ($hasPositiveExperience) { '明确正向用餐体验' } elseif ($hasRepeatedBehavior) { '明确重复到店行为' } else { '处于明确推荐清单' }
                [void]$result.Reasons.Add($reason)
            }
        }
    }
    if (-not $result.Recommended -and -not $result.Comparison -and -not $result.Negative) {
        $result.MentionOnly = $true
        [void]$result.Reasons.Add('只有店名提及，缺少推荐语义')
    }
    return $result
}

function New-Aggregate([string]$Name) {
    [pscustomobject]@{
        Name = $Name
        CommentKeys = New-StringSet
        Videos = New-StringSet
        Cities = New-StringSet
        Areas = New-StringSet
        Methods = New-StringSet
        RecommendationKeys = New-StringSet
        ComparisonKeys = New-StringSet
        NegativeKeys = New-StringSet
        MentionOnlyKeys = New-StringSet
        SemanticReasons = New-StringSet
        BestScore = -1
        EvidenceBv = ''
        EvidenceId = ''
        EvidenceText = ''
        MaxLikes = 0
    }
}

function Add-Hit($Aggregate, $Comment, [string]$Method, $Assessment) {
    $key = [string]$Comment.Key
    [void]$Aggregate.CommentKeys.Add($key)
    [void]$Aggregate.Videos.Add([string]$Comment.Bv)
    $source = $sourceByBvid[[string]$Comment.Bv]
    if ($null -ne $source) {
        if (-not [string]::IsNullOrWhiteSpace([string]$source.城市)) { [void]$Aggregate.Cities.Add([string]$source.城市) }
        if (-not [string]::IsNullOrWhiteSpace([string]$source.区县或县级市)) { [void]$Aggregate.Areas.Add([string]$source.区县或县级市) }
    }
    [void]$Aggregate.Methods.Add($Method)
    if ($Assessment.Recommended) { [void]$Aggregate.RecommendationKeys.Add($key) }
    if ($Assessment.Comparison) { [void]$Aggregate.ComparisonKeys.Add($key) }
    if ($Assessment.Negative) { [void]$Aggregate.NegativeKeys.Add($key) }
    if ($Assessment.MentionOnly) { [void]$Aggregate.MentionOnlyKeys.Add($key) }
    foreach ($reason in $Assessment.Reasons) { [void]$Aggregate.SemanticReasons.Add([string]$reason) }
    $likes = [int]$Comment.Likes
    if ($likes -gt $Aggregate.MaxLikes) { $Aggregate.MaxLikes = $likes }
    $score = $likes + $(if ($Assessment.Recommended) { 100000 } elseif ($Assessment.Negative) { 20 } elseif ($Assessment.Comparison) { -50 } else { 0 }) + $(if ($Method -eq '词典命中') { 50 } elseif ($Method -eq '推荐／店名句式') { 30 } else { 10 })
    if ($score -gt $Aggregate.BestScore) {
        $Aggregate.BestScore = $score
        $Aggregate.EvidenceBv = [string]$Comment.Bv
        $Aggregate.EvidenceId = [string]$Comment.Id
        $Aggregate.EvidenceText = Limit-Text ([string]$Comment.Text)
    }
}

if ($SemanticSelfTest) {
    $cases = @(
        @{ Text = '不如去京东食堂吃'; Name = '京东食堂'; Recommended = $false; Comparison = $true },
        @{ Text = '老牛不如淘米水'; Name = '淘米水'; Recommended = $false; Comparison = $true },
        @{ Text = '淘米水确实也不错'; Name = '淘米水'; Recommended = $true; Comparison = $false },
        @{ Text = '无锡人表示喜欢吃淘米水，那个红烧划水真的好吃'; Name = '淘米水'; Recommended = $true; Comparison = $false },
        @{ Text = '如果强制先吃那盘拼盘，那普通人还不如去吃汉巴味德'; Name = '汉巴味德'; Recommended = $false; Comparison = $true },
        @{ Text = '我推荐不如多走几步学前街卜岩面馆点一桌'; Name = '卜岩面馆'; Recommended = $true; Comparison = $false },
        @{ Text = '综合个人最推荐福乐，小笼的话无夕和超王记'; Name = '超王记'; Recommended = $true; Comparison = $false },
        @{ Text = '韩舍，五番亭，翼，一绪，这几个都可以'; Name = '五番亭'; Recommended = $true; Comparison = $false },
        @{ Text = '韩舍，五番亭，这几个都不推荐'; Name = '五番亭'; Recommended = $false; Comparison = $false }
    )
    foreach ($case in $cases) {
        $actual = Get-SemanticAssessment -Text $case.Text -Names @($case.Name)
        if ($actual.Recommended -ne $case.Recommended -or $actual.Comparison -ne $case.Comparison) {
            throw "语义自测失败：$($case.Text)；实际 Recommended=$($actual.Recommended), Comparison=$($actual.Comparison)"
        }
    }
    if ($null -ne (Clean-Candidate '京东食堂吃') -or $null -ne (Clean-Candidate '京东外卖更')) {
        throw '语义自测失败：京东食堂／外卖不应进入饭店候选。'
    }
    [pscustomobject]@{ 语义自测 = '通过'; 用例数 = $cases.Count } | Format-List
    return
}

if (-not (Test-Path -LiteralPath $RawDirectory -PathType Container)) {
    throw "原始评论目录不存在：$RawDirectory"
}
foreach ($required in @($AliasCsv, $EvidenceCsv, $VerificationCsv, $SourceCsv, $SemanticRuleCsv, $CandidateAuditCsv)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) { throw "缺少输入文件：$required" }
}

$dictionary = @(Import-Csv -LiteralPath $AliasCsv -Encoding UTF8)
$candidateAudit = @(Import-Csv -LiteralPath $CandidateAuditCsv -Encoding UTF8)
$acceptedAudit = @($candidateAudit | Where-Object { $_.审计结论 -in @('新增实体', '合并别名') })
$auditCityByCanonical = @{}
$auditAreaByCanonical = @{}

foreach ($group in @($acceptedAudit | Group-Object 标准店名)) {
    $cities = @($group.Group | ForEach-Object { [string]$_.城市 } | Where-Object { $_ } | Sort-Object -Unique)
    $areas = @($group.Group | ForEach-Object { [string]$_.区县或县级市 } | Where-Object { $_ } | Sort-Object -Unique)
    if ($cities.Count) { $auditCityByCanonical[[string]$group.Name] = $cities -join '／' }
    if ($areas.Count) { $auditAreaByCanonical[[string]$group.Name] = $areas -join '／' }
}

$augmentedDictionary = New-Object 'System.Collections.Generic.List[object]'
foreach ($entry in $dictionary) {
    $canonical = [string]$entry.标准店名
    $aliases = New-StringSet
    foreach ($alias in @(([string]$entry.别名列表) -split '\|' | Where-Object { $_ })) { [void]$aliases.Add([string]$alias) }
    foreach ($audit in @($acceptedAudit | Where-Object 标准店名 -eq $canonical)) { [void]$aliases.Add([string]$audit.候选写法) }
    [void]$augmentedDictionary.Add([pscustomobject]@{
        标准店名 = $canonical
        别名列表 = (@($aliases | Sort-Object) -join '|')
        精筛结论 = [string]$entry.精筛结论
        类别 = [string]$entry.类别
        备注 = [string]$entry.备注
    })
}
foreach ($auditGroup in @($acceptedAudit | Group-Object 标准店名)) {
    $canonical = [string]$auditGroup.Name
    if (@($dictionary | Where-Object 标准店名 -eq $canonical).Count) { continue }
    $aliases = New-StringSet
    [void]$aliases.Add($canonical)
    foreach ($audit in $auditGroup.Group) { [void]$aliases.Add([string]$audit.候选写法) }
    $firstAudit = $auditGroup.Group | Select-Object -First 1
    [void]$augmentedDictionary.Add([pscustomobject]@{
        标准店名 = $canonical
        别名列表 = (@($aliases | Sort-Object) -join '|')
        精筛结论 = '保留候选'
        类别 = [string]$firstAudit.类别
        备注 = "候选人工审计新增：$([string]$firstAudit.备注)"
    })
}
$dictionary = @($augmentedDictionary | ForEach-Object { $_ })
$verification = @(Import-Csv -LiteralPath $VerificationCsv -Encoding UTF8)
$semanticRuleByKey = @{}
foreach ($rule in @(Import-Csv -LiteralPath $SemanticRuleCsv -Encoding UTF8)) {
    $key = "{0}|{1}" -f ([string]$rule.标准店名).ToLowerInvariant(), [string]$rule.BV号
    $semanticRuleByKey[$key] = $rule
}
$sourceByBvid = @{}
foreach ($sourceRow in @(Import-Csv -LiteralPath $SourceCsv -Encoding UTF8)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$sourceRow.BV号)) {
        $sourceByBvid[[string]$sourceRow.BV号] = $sourceRow
    }
}
$semanticAuditByKey = @{}

function Apply-SemanticOverride($Assessment, [string]$Name, [string]$Bvid) {
    $key = "{0}|{1}" -f $Name.ToLowerInvariant(), $Bvid
    $wildcardKey = "{0}|*" -f $Name.ToLowerInvariant()
    $rule = if ($semanticRuleByKey.ContainsKey($key)) { $semanticRuleByKey[$key] } elseif ($semanticRuleByKey.ContainsKey($wildcardKey)) { $semanticRuleByKey[$wildcardKey] } else { $null }
    if ($null -eq $rule -or [string]$rule.语义策略 -ne '强制戏谑提及') { return $Assessment }
    $reasons = New-StringSet
    [void]$reasons.Add("人工语境复核：$([string]$rule.说明)")
    return [pscustomobject]@{
        Recommended = $false
        Comparison = $true
        Negative = $false
        MentionOnly = $false
        Reasons = $reasons
    }
}

function Add-SemanticAudit([string]$Canonical, [string]$Alias, $Comment, $Assessment) {
    $key = "{0}|{1}" -f $Canonical.ToLowerInvariant(), [string]$Comment.Key
    if (-not $semanticAuditByKey.ContainsKey($key)) {
        $source = $sourceByBvid[[string]$Comment.Bv]
        $semanticAuditByKey[$key] = [pscustomobject]@{
            Canonical = $Canonical
            Aliases = New-StringSet
            Comment = $Comment
            City = if ($auditCityByCanonical.ContainsKey($Canonical)) { [string]$auditCityByCanonical[$Canonical] } elseif ($null -eq $source) { '' } else { [string]$source.城市 }
            Area = if ($auditAreaByCanonical.ContainsKey($Canonical)) { [string]$auditAreaByCanonical[$Canonical] } elseif ($null -eq $source) { '' } else { [string]$source.区县或县级市 }
            Recommended = $false
            Comparison = $false
            Negative = $false
            MentionOnly = $false
            Reasons = New-StringSet
        }
    }
    $audit = $semanticAuditByKey[$key]
    [void]$audit.Aliases.Add($Alias)
    if ($Assessment.Recommended) { $audit.Recommended = $true }
    if ($Assessment.Comparison) { $audit.Comparison = $true }
    if ($Assessment.Negative) { $audit.Negative = $true }
    if ($Assessment.MentionOnly) { $audit.MentionOnly = $true }
    foreach ($reason in $Assessment.Reasons) { [void]$audit.Reasons.Add([string]$reason) }
}
$commentsByKey = @{}
$rawFileCount = 0
$skippedRawFileCount = 0

Get-ChildItem -LiteralPath $RawDirectory -Filter '*.json' -File | Sort-Object Name | ForEach-Object {
    if (-not $sourceByBvid.ContainsKey($_.BaseName)) {
        $skippedRawFileCount++
        return
    }
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
$cuePattern = '(?:推荐去|推荐吃|推荐|建议去|建议吃|建议|可以去|去吃|去尝尝|试试|首选|常去|必吃|安利|强推|店叫|叫做|叫)\s*(?<name>[A-Za-z0-9一-龥·&]{2,18})'
$beforeCuePattern = '(?<name>[A-Za-z0-9一-龥·&]{2,18})(?:很好吃|真好吃|好吃|很不错|确实不错|相当可以|值得去|值得吃)'
$suffixPattern = '[A-Za-z0-9一-龥·&]{2,18}(?:鸡鸭鹅大酒店|鸡鸭鹅大饭店|大酒店|大饭店|鸡汤面馆|馄饨店|酸辣汤|铁板烧|卤肉饭|酸菜鱼|猪脚饭|海棠糕|面馆|饭店|餐厅|食府|酒楼|饭庄|小馆|烤肉|烧烤|火锅|馄饨|小笼|包子|糕团|点心|炸鸡|拉面|生煎|小吃|酒店|菜馆)'

foreach ($comment in @($commentsByKey.Values)) {
    $text = ([string]$comment.Text -replace '\[[^\]]{1,20}\]', '') -replace '\s+', ' '
    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    $commentCandidates = @{}
    $semanticCache = @{}

    foreach ($aliasRow in $aliasRows) {
        if ($text.IndexOf($aliasRow.Alias, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $semanticKey = ([string]$aliasRow.Alias).ToLowerInvariant()
            if (-not $semanticCache.ContainsKey($semanticKey)) {
                $semanticCache[$semanticKey] = Get-SemanticAssessment -Text ([string]$comment.Text) -Names @([string]$aliasRow.Alias)
            }
            $assessment = Apply-SemanticOverride -Assessment $semanticCache[$semanticKey] -Name ([string]$aliasRow.Canonical) -Bvid ([string]$comment.Bv)
            Add-SemanticAudit -Canonical ([string]$aliasRow.Canonical) -Alias ([string]$aliasRow.Alias) -Comment $comment -Assessment $assessment
            Add-Hit $canonicalAggregates[$aliasRow.Canonical] $comment '词典命中' $assessment
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
            [void]$commentCandidates[$key].Methods.Add('推荐／店名句式')
        }
    }
    foreach ($match in [regex]::Matches($text, $beforeCuePattern, 'IgnoreCase')) {
        $name = Clean-Candidate $match.Groups['name'].Value
        if ($name) {
            $key = $name.ToLowerInvariant()
            if (-not $commentCandidates.ContainsKey($key)) { $commentCandidates[$key] = [pscustomobject]@{ Name = $name; Methods = New-StringSet } }
            [void]$commentCandidates[$key].Methods.Add('推荐／店名句式')
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
        if (-not $semanticCache.ContainsKey($key)) {
            $semanticCache[$key] = Get-SemanticAssessment -Text ([string]$comment.Text) -Names @([string]$candidate.Name)
        }
        $assessment = Apply-SemanticOverride -Assessment $semanticCache[$key] -Name ([string]$candidate.Name) -Bvid ([string]$comment.Bv)
        foreach ($method in $candidate.Methods) { Add-Hit $candidateAggregates[$key] $comment $method $assessment }
    }
}

$candidateRows = foreach ($aggregate in $candidateAggregates.Values) {
    $count = $aggregate.CommentKeys.Count
    $isDictionary = $aggregate.Methods.Contains('词典命中')
    $hasCue = $aggregate.Methods.Contains('推荐／店名句式')
    $recommendationCount = $aggregate.RecommendationKeys.Count
    $keep = $isDictionary -or $count -ge 2 -or ($hasCue -and $recommendationCount -gt 0 -and $aggregate.Name.Length -le 12)
    if (-not $keep) { continue }
    $level = if ($isDictionary) { 'A-词典命中' } elseif ($recommendationCount -gt 0) { 'B-存在明确推荐' } elseif ($count -ge 2) { 'C-多条提及' } else { 'D-单条宽松召回' }
    [pscustomobject]@{
        候选写法 = $aggregate.Name
        提及评论数 = $count
        涉及视频数 = $aggregate.Videos.Count
        涉及城市 = (($aggregate.Cities | Sort-Object) -join '／')
        涉及区县或县级市 = (($aggregate.Areas | Sort-Object) -join '／')
        有效推荐语境数 = $recommendationCount
        戏谑或比较语境数 = $aggregate.ComparisonKeys.Count
        负面语境数 = $aggregate.NegativeKeys.Count
        仅提及语境数 = $aggregate.MentionOnlyKeys.Count
        语义判定依据 = (($aggregate.SemanticReasons | Sort-Object) -join '；')
        最高点赞数 = $aggregate.MaxLikes
        召回方式 = (($aggregate.Methods | Sort-Object) -join '／')
        召回级别 = $level
        证据BV号 = $aggregate.EvidenceBv
        证据评论ID = $aggregate.EvidenceId
        证据摘录 = $aggregate.EvidenceText
        处理建议 = $(if ($isDictionary) { '进入实体精筛表' } elseif ($recommendationCount -gt 0) { '优先平台检索' } elseif ($aggregate.ComparisonKeys.Count -gt 0) { '只保留提及，不作为推荐' } else { '人工复核后再检索' })
    }
}
$candidateRows = @($candidateRows | Sort-Object @{ Expression = '有效推荐语境数'; Descending = $true }, @{ Expression = '提及评论数'; Descending = $true }, 候选写法)

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
        涉及城市 = $(if ($auditCityByCanonical.ContainsKey($canonical)) { [string]$auditCityByCanonical[$canonical] } else { (($aggregate.Cities | Sort-Object) -join '／') })
        涉及区县或县级市 = $(if ($auditAreaByCanonical.ContainsKey($canonical)) { [string]$auditAreaByCanonical[$canonical] } else { (($aggregate.Areas | Sort-Object) -join '／') })
        有效推荐语境数 = $aggregate.RecommendationKeys.Count
        戏谑或比较语境数 = $aggregate.ComparisonKeys.Count
        负面语境数 = $aggregate.NegativeKeys.Count
        仅提及语境数 = $aggregate.MentionOnlyKeys.Count
        推荐判定 = if ($aggregate.RecommendationKeys.Count -gt 0) { '存在明确推荐' } else { '未发现明确推荐' }
        语义判定依据 = (($aggregate.SemanticReasons | Sort-Object) -join '；')
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
$entityRows = @($entityRows | Sort-Object @{ Expression = '有效推荐语境数'; Descending = $true }, @{ Expression = '评论提及数'; Descending = $true }, 标准店名)

$semanticRows = foreach ($audit in $semanticAuditByKey.Values) {
    $roles = New-StringSet
    if ($audit.Recommended) { [void]$roles.Add('有效推荐') }
    if ($audit.Comparison) { [void]$roles.Add('戏谑或比较') }
    if ($audit.Negative) { [void]$roles.Add('负面评价') }
    if ($audit.MentionOnly) { [void]$roles.Add('仅提及') }
    [pscustomobject]@{
        标准店名 = $audit.Canonical
        命中写法 = (($audit.Aliases | Sort-Object) -join '／')
        BV号 = [string]$audit.Comment.Bv
        评论ID = [string]$audit.Comment.Id
        父评论ID = [string]$audit.Comment.ParentId
        发布时间 = [string]$audit.Comment.Published
        点赞数 = [int]$audit.Comment.Likes
        涉及城市 = $audit.City
        涉及区县或县级市 = $audit.Area
        语义角色 = (($roles | Sort-Object) -join '／')
        是否有效推荐 = if ($audit.Recommended) { '是' } else { '否' }
        判定依据 = (($audit.Reasons | Sort-Object) -join '；')
        评论文本 = Remove-UserReferences ([string]$audit.Comment.Text)
        数据来源 = [string]$audit.Comment.SourceType
    }
}
$semanticRows = @($semanticRows | Sort-Object @{ Expression = '是否有效推荐'; Descending = $true }, 标准店名, BV号, 评论ID)
$recommendedRows = @($semanticRows | Where-Object { $_.是否有效推荐 -eq '是' })

$candidateParent = Split-Path -Parent $CandidateOutput
$entityParent = Split-Path -Parent $EntityOutput
$semanticParent = Split-Path -Parent $SemanticOutput
$recommendedParent = Split-Path -Parent $RecommendedOutput
New-Item -ItemType Directory -Force -Path $candidateParent, $entityParent, $semanticParent, $recommendedParent | Out-Null
$candidateRows | Export-Csv -LiteralPath $CandidateOutput -NoTypeInformation -Encoding UTF8
$entityRows | Export-Csv -LiteralPath $EntityOutput -NoTypeInformation -Encoding UTF8
$semanticRows | Export-Csv -LiteralPath $SemanticOutput -NoTypeInformation -Encoding UTF8
$recommendedRows | Export-Csv -LiteralPath $RecommendedOutput -NoTypeInformation -Encoding UTF8

[pscustomobject]@{
    原始JSON文件数 = $rawFileCount
    排除非本批次JSON文件数 = $skippedRawFileCount
    原始评论数 = @($commentsByKey.Values | Where-Object { $_.SourceType -eq '原始评论' }).Count
    新增历史精选证据数 = $selectedEvidenceAdded
    实际可扫描记录数 = $commentsByKey.Count
    宽松候选写法数 = $candidateRows.Count
    精筛实体数 = $entityRows.Count
    实体评论语义标注数 = $semanticRows.Count
    有效推荐评论数 = $recommendedRows.Count
    候选输出 = (Resolve-Path -LiteralPath $CandidateOutput).Path
    精筛输出 = (Resolve-Path -LiteralPath $EntityOutput).Path
    语义标注输出 = (Resolve-Path -LiteralPath $SemanticOutput).Path
    有效推荐输出 = (Resolve-Path -LiteralPath $RecommendedOutput).Path
} | Format-List
