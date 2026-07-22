[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$Bvid,

    [string]$OutputDirectory = (Join-Path $env:TEMP 'healthy_life_bili_comments'),

    [ValidateRange(250, 5000)]
    [int]$MinDelayMs = 350
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$null = New-Item -ItemType Directory -Force -Path $OutputDirectory
$requestHeaders = @{
    Referer = 'https://www.bilibili.com/'
    'User-Agent' = 'Mozilla/5.0'
}
$jsonSerializer = $null

function Invoke-BiliApi {
    param([Parameter(Mandatory = $true)][string]$Uri)

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        Start-Sleep -Milliseconds $MinDelayMs
        try {
            $response = Invoke-RestMethod -Uri $Uri -Headers $requestHeaders -TimeoutSec 30
            # Windows PowerShell 5.1 may leave unusually large JSON responses as
            # strings even when Bilibili sends the correct application/json type.
            if ($response -is [string]) {
                try {
                    $response = $response | ConvertFrom-Json
                }
                catch {
                    # ConvertFrom-Json treats property names case-insensitively and
                    # rejects real comments containing keys such as Apple Pay and
                    # apple pay. JavaScriptSerializer preserves both spellings.
                    if ($null -eq $jsonSerializer) {
                        Add-Type -AssemblyName System.Web.Extensions
                        $jsonSerializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
                        $jsonSerializer.MaxJsonLength = [int]::MaxValue
                    }
                    $response = $jsonSerializer.DeserializeObject($response)
                }
            }
            $hasCode = $false
            $responseCode = $null
            if ($response -is [System.Collections.IDictionary]) {
                $hasCode = $response.ContainsKey('code')
                if ($hasCode) { $responseCode = $response['code'] }
            }
            else {
                $codeProperty = $response.PSObject.Properties['code']
                $hasCode = $null -ne $codeProperty
                if ($hasCode) { $responseCode = $codeProperty.Value }
            }
            if (-not $hasCode) {
                $responseType = if ($null -eq $response) { 'null' } else { $response.GetType().FullName }
                throw "Bilibili API response did not include code ($responseType): $Uri"
            }
            if ($responseCode -ne 0) {
                throw "Bilibili API code $($response.code): $($response.message)"
            }
            return $response
        }
        catch {
            if ($_.Exception.Message -match '\b412\b') {
                throw
            }
            if ($attempt -eq 3) {
                throw
            }
            Start-Sleep -Seconds $attempt
        }
    }
}

function Get-WbiMixinKey {
    # Anonymous nav responses may use code -101 while still returning the public
    # WBI image keys needed by the web client, so validate the keys themselves.
    Start-Sleep -Milliseconds $MinDelayMs
    $nav = Invoke-RestMethod -Uri 'https://api.bilibili.com/x/web-interface/nav' -Headers $requestHeaders -TimeoutSec 30
    if ($null -eq $nav.data.wbi_img) {
        throw "Bilibili nav response did not include WBI keys: $($nav.code) $($nav.message)"
    }
    $imageKey = [IO.Path]::GetFileNameWithoutExtension(([uri]$nav.data.wbi_img.img_url).AbsolutePath)
    $subKey = [IO.Path]::GetFileNameWithoutExtension(([uri]$nav.data.wbi_img.sub_url).AbsolutePath)
    $mixinTable = @(
        46, 47, 18, 2, 53, 8, 23, 32, 15, 50, 10, 31, 58, 3, 45, 35,
        27, 43, 5, 49, 33, 9, 42, 19, 29, 28, 14, 39, 12, 38, 41, 13,
        37, 48, 7, 16, 24, 55, 40, 61, 26, 17, 0, 1, 60, 51, 30, 4,
        22, 25, 54, 21, 56, 59, 6, 63, 57, 62, 11, 36, 20, 34, 44, 52
    )
    $rawKey = $imageKey + $subKey
    return (-join ($mixinTable | ForEach-Object { $rawKey[$_] })).Substring(0, 32)
}

function New-WbiUri {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUri,
        [Parameter(Mandatory = $true)]$Parameters,
        [Parameter(Mandatory = $true)][string]$MixinKey
    )

    $Parameters['wts'] = [string][DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $parts = foreach ($key in ($Parameters.Keys | Sort-Object)) {
        $value = ([string]$Parameters[$key]) -replace "[!'()*]", ''
        [uri]::EscapeDataString([string]$key) + '=' + [uri]::EscapeDataString($value)
    }
    $query = $parts -join '&'
    $md5 = [Security.Cryptography.MD5]::Create()
    try {
        $hash = $md5.ComputeHash([Text.Encoding]::UTF8.GetBytes($query + $MixinKey))
        $signature = -join ($hash | ForEach-Object { $_.ToString('x2') })
    }
    finally {
        $md5.Dispose()
    }
    return "$BaseUri`?$query&w_rid=$signature"
}

function Add-MinimalReply {
    param(
        [Parameter(Mandatory = $true)]$Reply,
        [Parameter(Mandatory = $true)][string]$VideoBvid,
        [Parameter(Mandatory = $true)]$ReplyTable
    )

    if ($null -eq $Reply -or $null -eq $Reply.rpid) {
        return
    }

    $replyId = [string]$Reply.rpid
    if ($ReplyTable.ContainsKey($replyId)) {
        return
    }

    $ReplyTable[$replyId] = [pscustomobject]@{
        BV号       = $VideoBvid
        评论ID     = $replyId
        父评论ID   = if ([long]$Reply.parent -eq 0) { '' } else { [string]$Reply.parent }
        发布时间   = [DateTimeOffset]::FromUnixTimeSeconds([long]$Reply.ctime).ToOffset([TimeSpan]::FromHours(8)).ToString('yyyy-MM-dd HH:mm:ss')
        点赞数     = [long]$Reply.like
        评论文本   = [string]$Reply.content.message
    }
}

foreach ($videoBvid in $Bvid) {
    $startedAt = Get-Date
    try {
        $detailUrl = "https://api.bilibili.com/x/web-interface/view/detail?bvid=$videoBvid"
        $detail = Invoke-BiliApi -Uri $detailUrl
        $view = $detail.data.View
        $aid = [long]$view.aid
        $advertisedCount = [int]$view.stat.reply

        $replyTable = [System.Collections.Generic.Dictionary[string, object]]::new()
        $rootTable = [System.Collections.Generic.Dictionary[string, object]]::new()
        $mixinKey = Get-WbiMixinKey
        $nextOffset = ''
        $topLevelPages = 0

        do {
            $pagination = @{ offset = $nextOffset } | ConvertTo-Json -Compress
            $mainUrl = New-WbiUri -BaseUri 'https://api.bilibili.com/x/v2/reply/wbi/main' -MixinKey $mixinKey -Parameters ([ordered]@{
                mode = '3'
                oid = [string]$aid
                pagination_str = $pagination
                plat = '1'
                seek_rpid = ''
                type = '1'
                web_location = '1315875'
            })
            $main = Invoke-BiliApi -Uri $mainUrl
            $pageReplies = @($main.data.replies | Where-Object { $null -ne $_ -and $null -ne $_.rpid })
            if ($pageReplies.Count -eq 0) {
                break
            }
            $topLevelPages++

            foreach ($root in $pageReplies) {
                if ($null -eq $root -or $null -eq $root.rpid) {
                    continue
                }
                Add-MinimalReply -Reply $root -VideoBvid $videoBvid -ReplyTable $replyTable
                $rootId = [string]$root.rpid
                if (-not $rootTable.ContainsKey($rootId)) {
                    $rootTable[$rootId] = $root
                }
                if ($null -ne $root.PSObject.Properties['replies']) {
                    foreach ($inlineChild in @($root.replies | Where-Object { $null -ne $_ })) {
                        Add-MinimalReply -Reply $inlineChild -VideoBvid $videoBvid -ReplyTable $replyTable
                    }
                }
            }

            $isEnd = [bool]$main.data.cursor.is_end
            $nextOffset = ''
            if (-not $isEnd -and $null -ne $main.data.cursor.pagination_reply) {
                $nextOffset = [string]$main.data.cursor.pagination_reply.next_offset
            }
            if ($topLevelPages -gt 2000) {
                throw "Top-level pagination exceeded safety limit for $videoBvid"
            }
        } while (-not $isEnd -and -not [string]::IsNullOrWhiteSpace($nextOffset))

        $replyPages = 0
        foreach ($root in $rootTable.Values) {
            if ([int]$root.rcount -le 0) {
                continue
            }

            $inlineReplyCount = 0
            if ($null -ne $root.PSObject.Properties['replies']) {
                $inlineReplyCount = @($root.replies | Where-Object { $null -ne $_ }).Count
            }
            if ($inlineReplyCount -ge [int]$root.rcount) {
                continue
            }

            $pageNumber = 1
            do {
                $childUrl = "https://api.bilibili.com/x/v2/reply/reply?type=1&oid=$aid&root=$($root.rpid)&pn=$pageNumber&ps=20"
                $children = Invoke-BiliApi -Uri $childUrl
                $replyPages++
                $pageReplies = @($children.data.replies | Where-Object { $null -ne $_ -and $null -ne $_.rpid })
                foreach ($child in $pageReplies) {
                    Add-MinimalReply -Reply $child -VideoBvid $videoBvid -ReplyTable $replyTable
                }

                $pageCount = [int]$children.data.page.count
                $pageSize = [int]$children.data.page.size
                $hasMore = ($pageNumber * $pageSize) -lt $pageCount -and $pageReplies.Count -gt 0
                $pageNumber++
            } while ($hasMore)
        }

        $comments = @($replyTable.Values | Sort-Object 发布时间, 评论ID)
        $output = [pscustomobject]@{
            BV号           = $videoBvid
            视频标题       = [string]$view.title
            AID            = $aid
            页面评论总数   = $advertisedCount
            实际抓取数     = $comments.Count
            顶层评论数     = $rootTable.Count
            回复数         = $comments.Count - $rootTable.Count
            顶层分页数     = $topLevelPages
            回复分页数     = $replyPages
            完整           = $comments.Count -eq $advertisedCount
            采集时间       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            评论           = $comments
        }

        $outputPath = Join-Path $OutputDirectory "$videoBvid.json"
        $output | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -Path $outputPath

        [pscustomobject]@{
            BV号         = $videoBvid
            页面评论总数 = $advertisedCount
            实际抓取数   = $comments.Count
            顶层评论数   = $rootTable.Count
            回复数       = $comments.Count - $rootTable.Count
            完整         = $comments.Count -eq $advertisedCount
            耗时秒       = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 1)
            输出文件     = $outputPath
            错误         = ''
        } | ConvertTo-Json -Compress
    }
    catch {
        [pscustomobject]@{
            BV号         = $videoBvid
            页面评论总数 = ''
            实际抓取数   = ''
            顶层评论数   = ''
            回复数       = ''
            完整         = $false
            耗时秒       = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 1)
            输出文件     = ''
            错误         = $_.Exception.Message
        } | ConvertTo-Json -Compress

        if ($_.Exception.Message -match '\b412\b') {
            throw
        }
    }
}
