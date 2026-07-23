[CmdletBinding()]
param(
    [string]$DataDirectory = '',
    [string]$OutputMarkdown = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$dietRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($DataDirectory)) { $DataDirectory = Join-Path $dietRoot '饭店数据' }
if ([string]::IsNullOrWhiteSpace($OutputMarkdown)) { $OutputMarkdown = Join-Path $dietRoot '江浙本地饭店.md' }

$entityCsv = Join-Path $DataDirectory '饭店实体精筛.csv'
$aliasCsv = Join-Path $DataDirectory '饭店别名词典.csv'
$verificationCsv = Join-Path $DataDirectory '门店核验.csv'
$evidenceCsv = Join-Path $DataDirectory '评论推荐.csv'
$semanticCsv = Join-Path $DataDirectory '饭店推荐语义标注.csv'
foreach ($path in @($entityCsv, $aliasCsv, $verificationCsv, $evidenceCsv, $semanticCsv, $OutputMarkdown)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "缺少输入文件：$path" }
}

function Escape-Markdown([string]$Value, [string]$Fallback = '待核验') {
    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq '未获取') { return $Fallback }
    return (($Value -replace '\s+', ' ').Trim() -replace '\|', '／')
}

function Get-NameSet($Entity, $AliasRow) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    [void]$set.Add([string]$Entity.标准店名)
    if ($null -ne $AliasRow) {
        foreach ($name in @(([string]$AliasRow.别名列表) -split '\|' | Where-Object { $_ })) { [void]$set.Add([string]$name) }
    }
    return $set
}

function Convert-ReviewCount([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq '未获取') { return -1L }
    $normalized = $Value.Trim().Replace(',', '')
    if ($normalized -match '^(\d+(?:\.\d+)?)万\+$') { return [long]([double]$matches[1] * 10000) }
    if ($normalized -match '^\d+$') { return [long]$normalized }
    return -1L
}

$entities = @(Import-Csv -LiteralPath $entityCsv -Encoding UTF8)
$aliases = @(Import-Csv -LiteralPath $aliasCsv -Encoding UTF8)
$verification = @(Import-Csv -LiteralPath $verificationCsv -Encoding UTF8)
$evidence = @(Import-Csv -LiteralPath $evidenceCsv -Encoding UTF8)
$semantic = @(Import-Csv -LiteralPath $semanticCsv -Encoding UTF8)
$dishVocabulary = @(
    '开洋三鲜馄饨', '开洋馄饨', '虾仁小馄饨', '三鲜馄饨', '小馄饨', '馄饨',
    '松子牛肉年糕', '松子牛肉', '无锡酱排骨', '无锡排骨', '酱排骨', '红烧肉', '蹄髈', '筒肠', '同肠',
    '肉酿面筋', '什锦面筋', '菜饭', '小笼包', '小笼馒头', '小笼', '汤包', '蟹黄汤包',
    '大排面', '大肠面', '鳝丝面', '鸡汤面', '春笋面', '阳春面', '拌面', '炒浇面', '面',
    '玉兰饼', '梅花糕', '海棠糕', '糖芋头', '糕团', '油酥', '油赞子', '生煎', '锅贴', '烧卖',
    '肉包', '菜包', '豆沙包', '包子', '酸辣汤', '香酥鸭', '梁溪脆鳝', '红烧划水', '醉虾',
    '热米皮', '肉夹馍', '卤肉饭', '拉面', '酸菜鱼', '猪脚汤', '鸭血粉丝', '大麦粥',
    '烤肉', '牛排', '炸鸡', '火锅', '羊肉面', '蟹粉', '点心'
)
$rows = @()

foreach ($entity in $entities) {
    $aliasRow = $aliases | Where-Object 标准店名 -eq $entity.标准店名 | Select-Object -First 1
    $names = Get-NameSet -Entity $entity -AliasRow $aliasRow
    $verifiedRows = @($verification | Where-Object { $names.Contains([string]$_.标准店名) })
    # 同名多分店时，只采用点评评价数最高且有评分的门店；没有评分时才退回至评价数最高的已核验门店。
    $ratedRows = @($verifiedRows | Where-Object {
        -not [string]::IsNullOrWhiteSpace([string]$_.点评评分) -and [string]$_.点评评分 -ne '未获取'
    })
    $displayVerification = @(
        $(if ($ratedRows.Count) { $ratedRows } else { $verifiedRows }) |
            Sort-Object @{ Expression = { Convert-ReviewCount ([string]$_.点评评价数) }; Descending = $true },
                        @{ Expression = { [string]$_.点评核验日 }; Descending = $true } |
            Select-Object -First 1
    )
    $historicalEvidence = @($evidence | Where-Object { $names.Contains([string]$_.推荐店名) })
    $semanticRows = @($semantic | Where-Object 标准店名 -eq $entity.标准店名)
    $recommendedRows = @($semanticRows | Where-Object 是否有效推荐 -eq '是')
    $dishSourceRow = $recommendedRows | Sort-Object @{ Expression = { [int]$_.点赞数 }; Descending = $true }, @{ Expression = { ([string]$_.评论文本).Length }; Ascending = $true } | Select-Object -First 1

    $dishSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($value in @($displayVerification | ForEach-Object { $_.点评推荐菜 }) + @($historicalEvidence | ForEach-Object { $_.推荐菜 })) {
        if ([string]::IsNullOrWhiteSpace([string]$value) -or [string]$value -eq '未获取') { continue }
        foreach ($dish in @(([string]$value) -split '[／/、|]' | Where-Object { $_ })) { [void]$dishSet.Add(([string]$dish).Trim()) }
    }
    foreach ($commentRow in @($dishSourceRow)) {
        if ($null -eq $commentRow) { continue }
        $commentTextForDish = [string]$commentRow.评论文本
        foreach ($dish in $dishVocabulary) {
            if ($commentTextForDish.Contains($dish)) { [void]$dishSet.Add($dish) }
        }
    }
    $prunedDishes = New-Object 'System.Collections.Generic.List[string]'
    foreach ($dish in @($dishSet | Sort-Object @{ Expression = { ([string]$_).Length }; Descending = $true }, @{ Expression = { [string]$_ } })) {
        if (@($prunedDishes | Where-Object { ([string]$_).Contains([string]$dish) }).Count -eq 0) { [void]$prunedDishes.Add([string]$dish) }
        if ($prunedDishes.Count -ge 8) { break }
    }
    $dishes = if ($prunedDishes.Count) { ($prunedDishes -join '／') } else { '评论未明确／待人工补充' }

    $scores = @($displayVerification | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace([string]$_.点评评分) -and [string]$_.点评评分 -ne '未获取') {
            $branch = if ([string]::IsNullOrWhiteSpace([string]$_.分店) -or [string]$_.分店 -eq '分店待确认') { '' } else { "$($_.分店)：" }
            "$branch$($_.点评评分)"
        }
    } | Where-Object { $_ } | Sort-Object -Unique)
    $rating = if ($scores.Count) { $scores -join '／' } else { '待平台核验' }

    $perCapitaValues = @($displayVerification | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace([string]$_.点评人均) -and [string]$_.点评人均 -ne '未获取') {
            $branch = if ([string]::IsNullOrWhiteSpace([string]$_.分店) -or [string]$_.分店 -eq '分店待确认') { '' } else { "$($_.分店)：" }
            "$branch 点评¥$($_.点评人均)".Trim()
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$_.高德人均) -and [string]$_.高德人均 -ne '未获取') {
            $branch = if ([string]::IsNullOrWhiteSpace([string]$_.分店) -or [string]$_.分店 -eq '分店待确认') { '' } else { "$($_.分店)：" }
            "$branch 高德¥$($_.高德人均)".Trim()
        }
    } | Where-Object { $_ } | Sort-Object -Unique)
    $perCapita = if ($perCapitaValues.Count) { $perCapitaValues -join '／' } else { '待平台核验' }

    $locations = @($displayVerification | ForEach-Object {
        $parts = @()
        if (-not [string]::IsNullOrWhiteSpace([string]$_.分店) -and [string]$_.分店 -ne '分店待确认') { $parts += [string]$_.分店 }
        if (-not [string]::IsNullOrWhiteSpace([string]$_.区县)) { $parts += [string]$_.区县 }
        if (-not [string]::IsNullOrWhiteSpace([string]$_.地址) -and [string]$_.地址 -ne '待确认' -and [string]$_.地址 -ne '分店待确认') { $parts += [string]$_.地址 }
        if ($parts.Count) { $parts -join '·' }
    } | Where-Object { $_ } | Sort-Object -Unique)
    $city = if ([string]::IsNullOrWhiteSpace([string]$entity.涉及城市)) { '无锡市' } else { (([string]$entity.涉及城市) -split '／')[0] }
    if ($locations.Count) {
        $location = $locations -join '<br>'
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$entity.涉及区县或县级市)) {
        $location = "$($entity.涉及区县或县级市)（分店待确认）"
    }
    else {
        $location = "$city（历史评论来源；分店待确认）"
    }

    $recommendationCount = [int]$entity.有效推荐语境数
    $isAmbiguous = [string]$entity.精筛结论 -eq '待歧义消解'
    $hasConcreteLocation = @($verifiedRows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.地址) -and [string]$_.地址 -notin @('待确认', '分店待确认') }).Count -gt 0
    $hasPlatformEvidence = [string]$entity.平台匹配状态 -ne '待平台核验'
    if ($recommendationCount -eq 0) {
        $confidence = '待确认（0 条）'
        $confidenceRank = 0
    }
    elseif ($isAmbiguous) {
        $confidence = "低（$recommendationCount 条）"
        $confidenceRank = 1
    }
    elseif ($recommendationCount -ge 3 -and $hasConcreteLocation -and $hasPlatformEvidence) {
        $confidence = "高（$recommendationCount 条）"
        $confidenceRank = 3
    }
    elseif ($recommendationCount -ge 2 -or ($recommendationCount -ge 1 -and $hasConcreteLocation)) {
        $confidence = "中（$recommendationCount 条）"
        $confidenceRank = 2
    }
    else {
        $confidence = "低（$recommendationCount 条）"
        $confidenceRank = 1
    }

    $chosenComment = $dishSourceRow
    if ($null -eq $chosenComment) {
        $chosenComment = $semanticRows | Sort-Object @{ Expression = { [int]$_.点赞数 }; Descending = $true }, @{ Expression = { ([string]$_.评论文本).Length }; Ascending = $true } | Select-Object -First 1
    }
    if ($null -eq $chosenComment) {
        $commentText = '暂无可追溯评论原文'
    }
    else {
        $prefix = if ([string]$chosenComment.是否有效推荐 -eq '是') { '' } else { "【$($chosenComment.语义角色)，非有效推荐】" }
        $commentText = "$prefix[$($chosenComment.BV号)](https://www.bilibili.com/video/$($chosenComment.BV号)/) $($chosenComment.评论文本)"
    }

    $rows += [pscustomobject]@{
        City = $city
        Name = [string]$entity.标准店名
        Dishes = $dishes
        Rating = $rating
        PerCapita = $perCapita
        Location = $location
        Confidence = $confidence
        ConfidenceRank = $confidenceRank
        RecommendationCount = $recommendationCount
        Comment = $commentText
    }
}

$lines = New-Object 'System.Collections.Generic.List[string]'
[void]$lines.Add('## 人工确认总表（按地级市）')
[void]$lines.Add('')
[void]$lines.Add('表内“评分／人均”只填写已取得的平台值；“待平台核验”表示当前没有可靠数据。同名多分店均有记录时，评分、人均、位置和推荐菜只取点评评价数最多且有评分的门店。置信度用于安排人工复核优先级：高＝至少 3 条有效推荐且已有平台位置，中＝至少 2 条推荐或已有具体位置，低＝仅 1 条推荐或名称有歧义，待确认＝没有有效推荐；括号仅显示有效推荐条数。评论原文优先选择点赞较高的有效推荐；没有有效推荐时保留一条非推荐原文用于排除判断。')
[void]$lines.Add('')
foreach ($cityName in @('无锡市', '苏州市', '常州市', '镇江市', '南京市', '泰州市')) {
    $cityRows = @($rows | Where-Object City -eq $cityName | Sort-Object @{ Expression = 'ConfidenceRank'; Descending = $true }, @{ Expression = 'RecommendationCount'; Descending = $true }, Name)
    [void]$lines.Add("### $cityName（$($cityRows.Count) 家）")
    [void]$lines.Add('')
    [void]$lines.Add('| 饭店名 | 推荐菜 | 评分 | 人均消费 | 位置 | 置信度 | 评论原文 |')
    [void]$lines.Add('| --- | --- | --- | --- | --- | --- | --- |')
    if (-not $cityRows.Count) {
        [void]$lines.Add('| 暂无可确认实体 | — | — | — | — | 待审计 | 当前精筛结果暂无该市实体；请复核评论采集状态和候选人工审计表。 |')
    }
    else {
        foreach ($row in $cityRows) {
            $cells = @(
                (Escape-Markdown $row.Name),
                (Escape-Markdown $row.Dishes '评论未明确／待人工补充'),
                (Escape-Markdown $row.Rating '待平台核验'),
                (Escape-Markdown $row.PerCapita '待平台核验'),
                (Escape-Markdown $row.Location),
                (Escape-Markdown $row.Confidence),
                (Escape-Markdown $row.Comment '暂无可追溯评论原文')
            )
            [void]$lines.Add('| ' + ($cells -join ' | ') + ' |')
        }
    }
    [void]$lines.Add('')
}

$beginMarker = '<!-- BEGIN AUTO:人工确认总表 -->'
$endMarker = '<!-- END AUTO:人工确认总表 -->'
$generatedBlock = $beginMarker + "`r`n" + ($lines -join "`r`n") + $endMarker
$markdown = Get-Content -LiteralPath $OutputMarkdown -Raw -Encoding UTF8
$pattern = '(?s)' + [regex]::Escape($beginMarker) + '.*?' + [regex]::Escape($endMarker)
if (-not [regex]::IsMatch($markdown, $pattern)) { throw '输出文档缺少人工确认总表生成标记。' }
$updated = [regex]::Replace($markdown, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($match) $generatedBlock }, 1)
$updated = $updated.TrimEnd("`r", "`n") + "`r`n"
[IO.File]::WriteAllText($OutputMarkdown, $updated, [Text.UTF8Encoding]::new($false))

[pscustomobject]@{
    输出文件 = (Resolve-Path -LiteralPath $OutputMarkdown).Path
    总实体数 = $rows.Count
    无锡市 = @($rows | Where-Object City -eq '无锡市').Count
    苏州市 = @($rows | Where-Object City -eq '苏州市').Count
    常州市 = @($rows | Where-Object City -eq '常州市').Count
    镇江市 = @($rows | Where-Object City -eq '镇江市').Count
    南京市 = @($rows | Where-Object City -eq '南京市').Count
    泰州市 = @($rows | Where-Object City -eq '泰州市').Count
} | Format-List
