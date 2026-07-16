#Requires -Version 5.1
<#
.SYNOPSIS
  Watch flgame.cloud Taiwan Jien silver market for new listings.
  Notification: Windows balloon tip (no accounts / bots / webhooks).

.EXAMPLE
  .\watch-silver.ps1 -Init
  .\watch-silver.ps1
  .\watch-silver.ps1 -Query "keyword"
  .\watch-silver.ps1 -DemoNotify
#>
[CmdletBinding()]
param(
    [string]$Market = "taiwan-jien",
    [string]$Category = "",
    [string]$Query = "",
    [switch]$Init,
    [switch]$AllPages,
    [switch]$DemoNotify,
    [string]$StateDir = ""
)

# 確保載入 System.Net.Http 組件（PowerShell 5.1 必備）
try { Add-Type -AssemblyName System.Net.Http } catch {}
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# 初始化狀態路徑
if (-not $StateDir) {
    if ($PSScriptRoot) { $StateDir = $PSScriptRoot }
    elseif ($MyInvocation.MyCommand.Path) { $StateDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    else { $StateDir = Join-Path $env:USERPROFILE "flgame-watch" }
}
$StatePath = Join-Path $StateDir "state.json"
$BaseUrl = "https://www.flgame.cloud/silver/"

# 全域共享 HttpClient 實例 (比 Invoke-WebRequest 快非常多)
if (-not $global:HTTP_CLIENT) {
    $handler = New-Object System.Net.Http.HttpClientHandler
    $global:HTTP_CLIENT = New-Object System.Net.Http.HttpClient -ArgumentList $handler
    $global:HTTP_CLIENT.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) flgame-watch/1.1")
    $global:HTTP_CLIENT.DefaultRequestHeaders.Add("Accept-Language", "zh-TW,zh;q=0.9")
}

function Show-Notify {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Body
    )

    Write-Host ""
    Write-Host "========== $Title ==========" -ForegroundColor Yellow
    Write-Host $Body -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Yellow
    Write-Host ""

    try {
        [Console]::Beep(880, 180)
        [Console]::Beep(1175, 220)
    } catch {}

    $icon = $null
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        
        $icon = New-Object System.Windows.Forms.NotifyIcon
        $icon.Icon = [System.Drawing.SystemIcons]::Information
        $icon.Visible = $true
        $icon.BalloonTipTitle = $Title
        
        $max = [Math]::Min(250, $Body.Length)
        $icon.BalloonTipText = $Body.Substring(0, $max)
        $icon.ShowBalloonTip(8000)
        
        # 讓通知有時間彈出，並釋放資源避免右下角圖示殘留
        Start-Sleep -Milliseconds 1500
    }
    catch {
        Write-Warning "Balloon tip failed: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $icon) {
            $icon.Visible = $false
            $icon.Dispose()
            # 強制垃圾回收，徹底清除右下角殘留圖示
            [System.GC]::Collect()
        }
    }
}

function Build-Url {
    param([int]$Page = 1)
    
    $params = [System.Collections.Generic.List[string]]::new()
    $params.Add("market=$([Uri]::EscapeDataString($Market))")
    if ($Category) { $params.Add("category=$([Uri]::EscapeDataString($Category))") }
    if ($Query)    { $params.Add("q=$([Uri]::EscapeDataString($Query))") }
    if ($Page -gt 1) { $params.Add("page=$Page") }
    
    return "${BaseUrl}?$($params -join '&')"
}

function Get-Html {
    param([string]$Url)
    $task = $global:HTTP_CLIENT.GetStringAsync($Url)
    return $task.GetAwaiter().GetResult()
}

function Get-UpdateTime {
    param([string]$Html)
    $regex = [regex]::new('([0-9]{4}-[0-9]{2}-[0-9]{2}\s+[0-9]{2}:[0-9]{2})', [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $match = $regex.Match($Html)
    if ($match.Success) {
        return $match.Value
    }
    return $null
}

function Get-TotalHint {
    param([string]$Html)
    $regex = [regex]::new('([0-9]+)-([0-9]+)\s*/\s*([0-9]+)', [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $m = $regex.Match($Html)
    if ($m.Success) {
        return @{
            From  = [int]$m.Groups[1].Value
            To    = [int]$m.Groups[2].Value
            Total = [int]$m.Groups[3].Value
        }
    }
    return $null
}

function Parse-Items {
    param([string]$Html)
    $pattern = '(?s)<tr>\s*<td>\s*<div class="silver-item-main">.*?<strong>(?<name>[^<]+)</strong>.*?<span class="category-badge">(?<cat>[^<]+)</span>.*?</td>\s*<td>(?<attrs>.*?)</td>\s*<td>(?<qty>[^<]*)</td>\s*<td class="price-cell">(?<price>[^<]*)</td>\s*<td>(?<stall>[^<]*)</td>\s*<td>(?<pos>[^<]*)</td>\s*</tr>'
    $regex = [regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $rowMatches = $regex.Matches($Html)
    
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $htmlCleanRegex = [regex]::new('<[^>]+>', [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $spaceRegex = [regex]::new('\s+', [System.Text.RegularExpressions.RegexOptions]::Compiled)

    foreach ($m in $rowMatches) {
        $attrsRaw = $m.Groups["attrs"].Value
        $attrText = $htmlCleanRegex.Replace($attrsRaw, ' ')
        $attrText = $spaceRegex.Replace($attrText, ' ').Trim()
        if (-not $attrText) { $attrText = "-" }

        $name  = $m.Groups["name"].Value.Trim()
        $cat   = $m.Groups["cat"].Value.Trim()
        $qty   = $m.Groups["qty"].Value.Trim()
        $price = $m.Groups["price"].Value.Trim()
        $stall = $m.Groups["stall"].Value.Trim()
        $pos   = $m.Groups["pos"].Value.Trim()
        $fp    = "$name|$cat|$attrText|$qty|$price|$stall|$pos"

        $result.Add([pscustomobject]@{
            Fingerprint = $fp
            Name        = $name
            Category    = $cat
            Attrs       = $attrText
            Qty         = $qty
            Price       = $price
            Stall       = $stall
            Position    = $pos
            # 用於保存該商品「首次被看見時」的官方更新時間
            FirstSeenSiteUpdate = "" 
        })
    }
    return @{ Items = $result.ToArray() }
}

function Fetch-Listings {
    $firstUrl = Build-Url -Page 1
    Write-Host "GET $firstUrl"
    $html = Get-Html -Url $firstUrl
    $updateTime = Get-UpdateTime -Html $html
    $hint = Get-TotalHint -Html $html
    $all = [System.Collections.Generic.List[PSCustomObject]]::new()
    $all.AddRange((Parse-Items -Html $html).Items)

    if ($AllPages -and $hint -and $hint.Total -gt $hint.To) {
        $pages = [int][Math]::Ceiling($hint.Total / 50.0)
        Write-Host ("Full scan: {0} items across ~{1} pages" -f $hint.Total, $pages)
        for ($p = 2; $p -le $pages; $p++) {
            Write-Host ("GET page {0}/{1}" -f $p, $pages)
            $pageHtml = Get-Html -Url (Build-Url -Page $p)
            $all.AddRange((Parse-Items -Html $pageHtml).Items)
            Start-Sleep -Milliseconds 300
        }
    }
    elseif ($hint) {
        Write-Host ("Parsed page 1: showing {0}-{1} of {2}" -f $hint.From, $hint.To, $hint.Total)
    }

    $snap = [ordered]@{
        FetchedAt   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        SiteUpdated = $updateTime
        Url         = $firstUrl
        Market      = $Market
        Category    = $Category
        Query       = $Query
        Count       = $all.Count
        Items       = $all.ToArray()
    }
    return $snap
}

function Save-State {
    param($Snapshot)
    $json = $Snapshot | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($StatePath, $json, [System.Text.Encoding]::UTF8)
    $upd = $Snapshot.SiteUpdated
    if (-not $upd) { $upd = "n/a" }
    Write-Host ("Saved state -> {0} ({1} items, site updated: {2})" -f $StatePath, $Snapshot.Count, $upd)
}

function Load-State {
    if (-not (Test-Path $StatePath)) { return $null }
    $json = [System.IO.File]::ReadAllText($StatePath, [System.Text.Encoding]::UTF8)
    return ($json | ConvertFrom-Json)
}

# 確保目錄存在
if (-not (Test-Path $StateDir)) {
    [System.IO.Directory]::CreateDirectory($StateDir) | Out-Null
}

# ==================== 1. 抓取與分析資料 ====================
$snap = Fetch-Listings
$updNow = $snap.SiteUpdated
if (-not $updNow) { $updNow = "(not found)" }
Write-Host ("Site update time: {0}" -f $updNow)

$prev = Load-State

# 建立舊商品的「首次看見時間」與「指紋」Map
$prevSeenMap = @{}
if ($prev -and $prev.Items) {
    foreach ($it in @($prev.Items)) {
        if ($it.Fingerprint) {
            # 兼容舊版 state.json：如果原本沒有 FirstSeenSiteUpdate，就預設為上次的更新時間
            $seenTime = $it.FirstSeenSiteUpdate
            if (-not $seenTime) { $seenTime = $prev.SiteUpdated }
            $prevSeenMap[$it.Fingerprint] = $seenTime
        }
    }
}

# 輔助：解析價格字串為純數字
function Parse-PriceInt ([string]$priceStr) {
    $cleaned = $priceStr -replace '[^\d]'
    if ($cleaned) { return [int]$cleaned }
    return 0
}

# 建立舊價格 Map，Key: "名稱|攤位|座標"
$prevPriceMap = @{}
$prevTime = ""
if ($prev) {
    $prevTime = $prev.FetchedAt
    if ($prev.Items) {
        foreach ($it in @($prev.Items)) {
            $key = "$($it.Name)|$($it.Stall)|$($it.Position)"
            $prevPriceMap[$key] = Parse-PriceInt -priceStr $it.Price
        }
    }
}

# 在記憶體中動態註記 [IsNew] 與 [降價資訊]
$hasNewItemForNotify = $false  # 用於判斷本次是否要彈窗通知
foreach ($it in $snap.Items) {
    $fp = [string]$it.Fingerprint
    
    # 判斷這個商品的首次出現時間
    if ($fp -and $prevSeenMap.ContainsKey($fp)) {
        # 以前看過了，沿用舊的首次看見時間
        $it.FirstSeenSiteUpdate = $prevSeenMap[$fp]
    } else {
        # 這是全新未見過的商品特徵！它的首次看見時間就是現在官方的 SiteUpdated
        $it.FirstSeenSiteUpdate = $updNow
        # 如果不是第一次初始化，且官方更新時間確實變了，就標記需要彈窗通知
        if ($prev) {
            $hasNewItemForNotify = $true
        }
    }

    # 決定網頁上是否要顯示 [NEW] 標籤：
    # 只要該商品首次被看見的時間，跟現在官方網頁的更新時間（$updNow）完全一致，它就是 NEW！
    if ($it.FirstSeenSiteUpdate -eq $updNow) {
        $it | Add-Member -MemberType NoteProperty -Name "IsNew" -Value $true -Force
    } else {
        $it | Add-Member -MemberType NoteProperty -Name "IsNew" -Value $false -Force
    }

    # 判斷降價 (同位置同攤位，且價格變便宜)
    $key = "$($it.Name)|$($it.Stall)|$($it.Position)"
    $newPrice = Parse-PriceInt -priceStr $it.Price
    if ($prevPriceMap.ContainsKey($key)) {
        $oldPrice = $prevPriceMap[$key]
        if ($newPrice -lt $oldPrice -and $newPrice -gt 0) {
            $it | Add-Member -MemberType NoteProperty -Name "PriceDropped" -Value $true -Force
            $it | Add-Member -MemberType NoteProperty -Name "OldPrice" -Value $oldPrice -Force
            $it | Add-Member -MemberType NoteProperty -Name "PriceDiff" -Value ($oldPrice - $newPrice) -Force
            $it | Add-Member -MemberType NoteProperty -Name "PrevTime" -Value $prevTime -Force
        } else {
            $it | Add-Member -MemberType NoteProperty -Name "PriceDropped" -Value $false -Force
        }
    } else {
        $it | Add-Member -MemberType NoteProperty -Name "PriceDropped" -Value $false -Force
    }
}

# ==================== 2. 自動產生 HTML 互動儀表板 ====================
$HtmlPath = Join-Path $StateDir "market-dashboard.html"
$StampPath = Join-Path $StateDir "live-stamp.js"
$IsFirstTime = -not (Test-Path $HtmlPath)

# 寫入時間戳記，供瀏覽器背景輪詢
$stampContent = "window.__last_market_update = `"$($snap.FetchedAt)`";"
[System.IO.File]::WriteAllText($StampPath, $stampContent, [System.Text.Encoding]::UTF8)

# 序列化商品資料
$itemsJson = $snap.Items | ConvertTo-Json -Depth 5 -Compress

# HTML 與 JS
$HtmlContent = @"
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>童話金銀商場儀表板</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
        h1 { margin-top: 0; color: #2c3e50; border-bottom: 2px solid #ecf0f1; padding-bottom: 15px; }
        .meta-info { font-size: 0.9em; color: #7f8c8d; margin-bottom: 20px; }
        .filter-section { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; border: 1px solid #e2e8f0; align-items: center; }
        .filter-group { display: flex; flex-direction: column; }
        .filter-group label { font-size: 0.85em; font-weight: bold; color: #4a5568; margin-bottom: 5px; }
        .filter-group input[type="text"], .filter-group input[type="number"] { padding: 8px 12px; border: 1px solid #cbd5e0; border-radius: 6px; font-size: 0.9em; outline: none; }
        .filter-group input:focus { border-color: #3182ce; box-shadow: 0 0 0 3px rgba(66,153,225,0.5); }
        
        /* Checkbox 樣式 */
        .checkbox-group { display: flex; align-items: center; height: 100%; padding-top: 15px; }
        .checkbox-group label { display: flex; align-items: center; font-size: 0.9em; font-weight: bold; color: #2d3748; cursor: pointer; user-select: none; }
        .checkbox-group input { width: 18px; height: 18px; margin-right: 8px; cursor: pointer; }

        table { width: 100%; border-collapse: collapse; margin-top: 15px; text-align: left; }
        th, td { padding: 12px 15px; border-bottom: 1px solid #e2e8f0; }
        th { background-color: #edf2f7; color: #2d3748; font-weight: bold; }
        tr:hover { background-color: #f7fafc; }
        .badge { background: #e2e8f0; padding: 3px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold; }
        .badge-attr { background: #ebf8ff; color: #2b6cb0; border: 1px solid #bee3f8; }
        .highlight { background-color: #fffaf0; }
        
        /* 亮點樣式 */
        .badge-new { background: #f56565; color: white; padding: 2px 6px; border-radius: 4px; font-size: 0.75em; margin-left: 8px; font-weight: bold; display: inline-block; vertical-align: middle; animation: pulse 2s infinite; }
        .price-drop { color: #48bb78; font-size: 0.8em; font-weight: bold; margin-top: 2px; }
        .price-stat { font-size: 0.75em; color: #718096; margin-top: 2px; }
        
        @keyframes pulse {
            0% { transform: scale(1); }
            50% { transform: scale(1.08); background-color: #e53e3e; }
            100% { transform: scale(1); }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>童話金銀商場儀表板</h1>
        <div class="meta-info">
            資料更新時間: $($snap.SiteUpdated) | 本地抓取時間: $($snap.FetchedAt) | 商品總數: $($snap.Count)
        </div>
        
        <div class="filter-section">
            <div class="filter-group">
                <label for="searchName">物品名稱關鍵字</label>
                <input type="text" id="searchName" placeholder="例如: 颶風斧" oninput="filterData()">
            </div>
            <div class="filter-group">
                <label for="searchAttr">裝備屬性篩選</label>
                <input type="text" id="searchAttr" placeholder="例如: 力量" oninput="filterData()">
            </div>
            <div class="filter-group">
                <label for="minAttrValue">屬性數值大於或等於 (>=)</label>
                <input type="number" id="minAttrValue" value="0" min="0" oninput="filterData()">
            </div>
            <div class="checkbox-group">
                <label>
                    <input type="checkbox" id="onlyNew" onchange="filterData()"> 只看全新上架 (NEW)
                </label>
            </div>
        </div>

        <div style="overflow-x: auto;">
            <table>
                <thead>
                    <tr>
                        <th>物品名稱</th>
                        <th>屬性/精鍊 (Attrs)</th>
                        <th>價格</th>
                        <th>數量</th>
                        <th>攤位名稱</th>
                        <th>座標位置</th>
                        <th>分類</th>
                    </tr>
                </thead>
                <tbody id="itemTableBody">
                    <!-- JS 渲染位置 -->
                </tbody>
            </table>
        </div>
    </div>

    <script>
        const rawData = $itemsJson;
        const pageLoadTime = "$($snap.FetchedAt)";

        // 1. 自動背景刷新檢查 (輪詢 JS 檔)
        function checkLiveUpdate() {
            const oldScript = document.getElementById('liveStampScript');
            if (oldScript) oldScript.remove();
            
            const script = document.createElement('script');
            script.id = 'liveStampScript';
            script.src = 'live-stamp.js?t=' + new Date().getTime();
            script.onload = function() {
                if (window.__last_market_update && window.__last_market_update !== pageLoadTime) {
                    console.log("市場資料已更新，重新載入頁面中...");
                    window.location.reload();
                }
            };
            document.head.appendChild(script);
        }
        setInterval(checkLiveUpdate, 5000); // 5 秒輪詢一次

        // 2. 價格統計：只篩選非裝備一般道具
        const isGeneralItem = function(item) {
            const cat = item.Category;
            return !cat.includes("武器") && !cat.includes("防具") && !cat.includes("裝備") && item.Attrs === "-";
        };

        const statsMap = {};
        rawData.forEach(function(item) {
            if (isGeneralItem(item)) {
                const priceNum = parseInt(item.Price.replace(/,/g, ''), 10) || 0;
                if (priceNum > 0) {
                    if (!statsMap[item.Name]) statsMap[item.Name] = [];
                    statsMap[item.Name].push(priceNum);
                }
            }
        });

        const itemStats = {};
        for (const name in statsMap) {
            const prices = statsMap[name];
            const min = Math.min.apply(null, prices);
            const max = Math.max.apply(null, prices);
            itemStats[name] = { min: min, max: max, count: prices.length };
        }

        // 數值縮減格式化
        function formatCompact(num) {
            if (num >= 10000) {
                const formatted = (num / 10000).toFixed(1).replace(/\.0$/, '');
                return formatted + '萬';
            }
            return num.toLocaleString();
        }

        // 3. 屬性過濾邏輯
        function parseAttributeValue(attrText, attrName) {
            if (!attrText || attrText === "-") return null;
            
            var query = attrName.trim();
            if (!query) return null;

            if (query === "精鍊" || query === "精煉") {
                var refineMatch = attrText.match(/\+([0-9]+)/);
                return refineMatch ? parseInt(refineMatch[1], 10) : null;
            }

            var regex = new RegExp(query + "\\s*\\+?\\s*([0-9]+)");
            var match = attrText.match(regex);
            return match ? parseInt(match[1], 10) : null;
        }

        // 4. 前端資料繪製
        function filterData() {
            var nameQuery = document.getElementById('searchName').value.toLowerCase();
            var attrQuery = document.getElementById('searchAttr').value.trim();
            var minVal = parseInt(document.getElementById('minAttrValue').value, 10) || 0;
            var onlyNew = document.getElementById('onlyNew').checked;
            
            var tbody = document.getElementById('itemTableBody');
            tbody.innerHTML = '';

            rawData.forEach(function(item) {
                var nameMatch = !nameQuery || item.Name.toLowerCase().indexOf(nameQuery) !== -1;
                
                var attrMatch = true;
                if (attrQuery) {
                    var extractedValue = parseAttributeValue(item.Attrs, attrQuery);
                    if (extractedValue !== null) {
                        attrMatch = extractedValue >= minVal;
                    } else {
                        attrMatch = false; 
                    }
                }

                // 「只看 NEW」過濾邏輯
                var newMatch = !onlyNew || item.IsNew;

                if (nameMatch && attrMatch && newMatch) {
                    var row = document.createElement('tr');
                    if (attrQuery) {
                        row.className = 'highlight';
                    }
                    
                    // NEW 標籤
                    var nameDisplay = '<strong>' + item.Name + '</strong>';
                    if (item.IsNew) {
                        nameDisplay += '<span class="badge-new">NEW</span>';
                    }
                    
                    // 價格資訊 / 降價 / 價格區間
                    var priceDisplay = '<div style="color: #e53e3e; font-weight: bold;">' + item.Price + '</div>';
                    if (item.PriceDropped) {
                        var tStr = "";
                        if (item.PrevTime) {
                            var parts = item.PrevTime.split(' ');
                            if (parts.length >= 2) {
                                var datePart = parts[0].substring(5); // 取 "MM-DD"
                                var timePart = parts[1].substring(0, 5); // 取 "HH:MM"
                                tStr = ' (對比 ' + datePart + ' ' + timePart + ')';
                            } else {
                                tStr = ' (對比上次)';
                            }
                        }
                        priceDisplay += '<div class="price-drop">↓ 便宜了 ' + item.PriceDiff.toLocaleString() + tStr + '</div>';
                    } else if (isGeneralItem(item) && itemStats[item.Name] && itemStats[item.Name].count > 1) {
                        var stat = itemStats[item.Name];
                        priceDisplay += '<div class="price-stat" title="市場最低價與最高價區間">區間: ' + 
                                        formatCompact(stat.min) + ' ~ ' + formatCompact(stat.max) + '</div>';
                    }
                    
                    var attrDisplay = item.Attrs !== "-" ? '<span class="badge badge-attr">' + item.Attrs + '</span>' : "-";
                    
                    row.innerHTML = '<td>' + nameDisplay + '</td>' +
                                    '<td>' + attrDisplay + '</td>' +
                                    '<td>' + priceDisplay + '</td>' +
                                    '<td>' + item.Qty + '</td>' +
                                    '<td>' + item.Stall + '</td>' +
                                    '<td style="color: #718096; font-size: 0.9em;">' + item.Position + '</td>' +
                                    '<td><span class="badge">' + item.Category + '</span></td>';
                    
                    tbody.appendChild(row);
                }
            });
        }

        // 初始化
        filterData();
    </script>
</body>
</html>
"@

[System.IO.File]::WriteAllText($HtmlPath, $HtmlContent, [System.Text.Encoding]::UTF8)
Write-Host "HTML Dashboard updated -> $HtmlPath" -ForegroundColor Green

# 只有在第一次建立或強制 Init 時，才主動打開瀏覽器
if ($IsFirstTime -or $Init) {
    try {
        Start-Process $HtmlPath
        Write-Host "Automatically opened dashboard." -ForegroundColor Cyan
    } catch {
        Write-Warning "Could not automatically open dashboard: $($_.Exception.Message)"
    }
} else {
    Write-Host "Dashboard updated in background. (Your browser tab will auto-refresh!)" -ForegroundColor Cyan
}

# ==================== 3. 初始化與 Demo 邏輯 ====================
if ($Init) {
    Save-State -Snapshot $snap
    Write-Host "Baseline initialized. Run again after the site update time changes."
    exit 0
}

if ($DemoNotify) {
    if (-not $prev) {
        Save-State -Snapshot $snap
        $prev = Load-State
    }
    # Demo 注入一個虛擬的新商品，其 FirstSeenSiteUpdate 強制設定為現在最新的時間，保證網頁一定判定為 NEW
    $fake = [pscustomobject]@{
        Fingerprint  = "DEMO|item|-|1|1|demo-stall|0, 0"
        Name         = "[DEMO] New item notify test"
        Category     = "item"
        Attrs        = "-"
        Qty          = "1"
        Price        = "1"
        Stall        = "demo-stall"
        Position     = "0, 0"
        FirstSeenSiteUpdate = $updNow
        IsNew        = $true
        PriceDropped = $false
    }
    $snap.Items = @($fake) + @($snap.Items)
    $snap.Count = @($snap.Items).Count
    $hasNewItemForNotify = $true
    Write-Host "Demo mode: injecting a fake new item to trigger notify." -ForegroundColor Magenta
}

if (-not $prev) {
    Save-State -Snapshot $snap
    Write-Host "No previous state. Baseline saved. Run again later to detect new items."
    exit 0
}

# ==================== 4. 新上架比對與通知邏輯 ====================
# 收集網頁上的「NEW」新商品
$newItems = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($it in @($snap.Items)) {
    if ($it.IsNew) {
        $newItems.Add($it)
    }
}

$prevUpdated = [string]$prev.SiteUpdated
$nowUpdated = [string]$snap.SiteUpdated
Write-Host ("Previous site update time: {0}" -f $(if ($prevUpdated) { $prevUpdated } else { "(unknown)" }))
Write-Host ("Current  site update time: {0}" -f $(if ($nowUpdated) { $nowUpdated } else { "(unknown)" }))

# 通知判定：
# 我們只有在「官方更新時間確實變了（或者 DemoMode）」且「確實有新增從未見過的 Fingerprint」時才彈窗
if (-not $hasNewItemForNotify) {
    Write-Host "No brand-new items to notify." -ForegroundColor Green
    if ($prevUpdated -and $nowUpdated -and ($prevUpdated -eq $nowUpdated)) {
        Write-Host "Site data timestamp unchanged -- wait for refresh." -ForegroundColor DarkYellow
    }
    Save-State -Snapshot $snap
    Write-Host "Done."
    exit 0
}

# 觸發 Windows 通知
$lines = [System.Collections.Generic.List[string]]::new()
# 篩選出本次最核心新上架的商品 (只取前 8 筆展示於通知)
foreach ($n in ($newItems | Select-Object -First 8)) {
    $attrDisplay = if ($n.Attrs -and $n.Attrs -ne "-") { " ($($n.Attrs))" } else { "" }
    $lines.Add("$($n.Name)$attrDisplay | $($n.Price) | $($n.Stall) @ $($n.Position)")
}
$body = "Found {0} new item(s) in update [{1}]`n{2}" -f $newItems.Count, $updNow, ($lines -join "`n")
Show-Notify -Title "童話金銀商場 - 新商品" -Body $body

Write-Host "New items in current site update [$updNow]:"
$newItems | Select-Object Name, Attrs, Price, Stall, Position, Qty, Category | Format-Table -AutoSize

if ($DemoNotify) {
    Write-Host "Demo done. State left unmodified."
} else {
    Save-State -Snapshot $snap
}

# ==================== 5. 自動同步至 GitHub Pages ====================
$EnableGithubSync = $true  # 是否啟用自動同步 (不想同步時可改為 $false)

if ($EnableGithubSync -and -not $DemoNotify) {
    Write-Host "正在同步最新資料至 GitHub Pages..." -ForegroundColor Cyan
    try {
        # 確保回到 $StateDir 目錄執行 Git 指令
        $currentDir = Get-Location
        Set-Location $StateDir
        
        # 檢查 Git 是否已初始化
        if (Test-Path (Join-Path $StateDir ".git")) {
            # 將網頁與狀態檔加入暫存區
            git add market-dashboard.html live-stamp.js state.json
            
            # Commit (使用目前時間作為備註)
            $commitMsg = "Auto-update market data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            git commit -m $commitMsg
            
            # 推送到 GitHub (請確保背景執行不會卡住)
            git push origin main --quiet
            
            Write-Host "GitHub Pages 同步成功！" -ForegroundColor Green
        } else {
            Write-Warning "偵測到該目錄尚未進行 git init，請先完成步驟三的設定。"
        }
    }
    catch {
        Write-Warning "GitHub 同步失敗: $($_.Exception.Message)"
    }
    finally {
        # 切回原本的工作目錄
        Set-Location $currentDir
    }
}
Write-Host "Done."
