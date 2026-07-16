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
    
    # 改用純字串拼接，不依賴 System.Web (啟動更快、跨平台相容)
    $params = [System.Collections.Generic.List[string]]::new()
    $params.Add("market=$([Uri]::EscapeDataString($Market))")
    if ($Category) { $params.Add("category=$([Uri]::EscapeDataString($Category))") }
    if ($Query)    { $params.Add("q=$([Uri]::EscapeDataString($Query))") }
    if ($Page -gt 1) { $params.Add("page=$Page") }
    
    # 使用 ${BaseUrl} 來確保 PowerShell 準確識別變數範圍
    return "${BaseUrl}?$($params -join '&')"
}

function Get-Html {
    param([string]$Url)
    # 使用 HttpClient 非同步取得網頁內容
    $task = $global:HTTP_CLIENT.GetStringAsync($Url)
    return $task.GetAwaiter().GetResult()
}

function Get-UpdateTime {
    param([string]$Html)
    # 優化：預先編譯 Regex 提升重複匹配效能
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
    # Regex 加上 Compiled 屬性，在處理大量資料時更快
    $pattern = '(?s)<tr>\s*<td>\s*<div class="silver-item-main">.*?<strong>(?<name>[^<]+)</strong>.*?<span class="category-badge">(?<cat>[^<]+)</span>.*?</td>\s*<td>(?<attrs>.*?)</td>\s*<td>(?<qty>[^<]*)</td>\s*<td class="price-cell">(?<price>[^<]*)</td>\s*<td>(?<stall>[^<]*)</td>\s*<td>(?<pos>[^<]*)</td>\s*</tr>'
    $regex = [regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $rowMatches = $regex.Matches($Html)
    
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    # 預先定義清理 HTML 標籤的 Regex
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
            Start-Sleep -Milliseconds 300 # HttpClient 較快，可以稍微縮短延遲
        }
    }
    elseif ($hint) {
        Write-Host ("Parsed page 1: showing {0}-{1} of {2}" -f $hint.From, $hint.To, $hint.Total)
        if ($hint.Total -gt $hint.To -and -not $Query) {
            Write-Host "Tip: use -AllPages for full market, or -Query keyword to watch specific items." -ForegroundColor DarkGray
        }
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
    # 使用深度確保所有 nested 屬性都有被序列化
    $json = $Snapshot | ConvertTo-Json -Depth 8
    # 強制使用無 BOM 的 UTF-8 寫入
    [System.IO.File]::WriteAllText($StatePath, $json, [System.Text.Encoding]::UTF8)
    
    $upd = $Snapshot.SiteUpdated
    if (-not $upd) { $upd = "n/a" }
    Write-Host ("Saved state -> {0} ({1} items, site updated: {2})" -f $StatePath, $Snapshot.Count, $upd)
}

function Load-State {
    if (-not (Test-Path $StatePath)) { return $null }
    # 使用 .NET 讀取效率更佳
    $json = [System.IO.File]::ReadAllText($StatePath, [System.Text.Encoding]::UTF8)
    return ($json | ConvertFrom-Json)
}

# 確保目錄存在
if (-not (Test-Path $StateDir)) {
    [System.IO.Directory]::CreateDirectory($StateDir) | Out-Null
}

# ==================== 1. 抓取資料 ====================
$snap = Fetch-Listings
$updNow = $snap.SiteUpdated
if (-not $updNow) { $updNow = "(not found)" }
Write-Host ("Site update time: {0}" -f $updNow)
Write-Host ("Fetched {0} listings at {1}" -f $snap.Count, $snap.FetchedAt)

# ==================== 2. 自動產生並開啟 HTML 互動儀表板 ====================
$HtmlPath = Join-Path $StateDir "market-dashboard.html"
$itemsJson = $snap.Items | ConvertTo-Json -Depth 5 -Compress

# 使用傳統 JS 字串拼接來徹底避免 PowerShell 與 JS 樣板字串（Backticks, $）的衝突
$HtmlContent = @"
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Taiwan Jien 銀幣市場監控</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; background-color: #f4f6f9; color: #333; margin: 0; padding: 20px; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
        h1 { margin-top: 0; color: #2c3e50; border-bottom: 2px solid #ecf0f1; padding-bottom: 15px; }
        .meta-info { font-size: 0.9em; color: #7f8c8d; margin-bottom: 20px; }
        .filter-section { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; border: 1px solid #e2e8f0; }
        .filter-group { display: flex; flex-direction: column; }
        .filter-group label { font-size: 0.85em; font-weight: bold; color: #4a5568; margin-bottom: 5px; }
        .filter-group input { padding: 8px 12px; border: 1px solid #cbd5e0; border-radius: 6px; font-size: 0.9em; outline: none; }
        .filter-group input:focus { border-color: #3182ce; box-shadow: 0 0 0 3px rgba(66,153,225,0.5); }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; text-align: left; }
        th, td { padding: 12px 15px; border-bottom: 1px solid #e2e8f0; }
        th { background-color: #edf2f7; color: #2d3748; font-weight: bold; }
        tr:hover { background-color: #f7fafc; }
        .badge { background: #e2e8f0; padding: 3px 8px; border-radius: 4px; font-size: 0.8em; font-weight: bold; }
        .badge-attr { background: #ebf8ff; color: #2b6cb0; border: 1px solid #bee3f8; }
        .highlight { background-color: #fffaf0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Taiwan Jien 銀幣市場監控儀表板</h1>
        <div class="meta-info">
            資料更新時間: $($snap.SiteUpdated) | 本地抓取時間: $($snap.FetchedAt) | 商品總數: $($snap.Count)
        </div>
        
        <div class="filter-section">
            <div class="filter-group">
                <label for="searchName">物品名稱關鍵字</label>
                <input type="text" id="searchName" placeholder="例如: 颶風斧" oninput="filterData()">
            </div>
            <div class="filter-group">
                <label for="searchAttr">裝備屬性篩選 (精鍊/力量/體質/物理攻擊)</label>
                <input type="text" id="searchAttr" placeholder="例如: 力量" oninput="filterData()">
            </div>
            <div class="filter-group">
                <label for="minAttrValue">屬性數值大於或等於 (>=)</label>
                <input type="number" id="minAttrValue" value="0" min="0" oninput="filterData()">
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
        // 載入資料 (PowerShell 會在此將 $itemsJson 轉譯為字串注入)
        const rawData = $itemsJson;

        // 從屬性文字中提取特定屬性的數值
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

        function filterData() {
            var nameQuery = document.getElementById('searchName').value.toLowerCase();
            var attrQuery = document.getElementById('searchAttr').value.trim();
            var minVal = parseInt(document.getElementById('minAttrValue').value, 10) || 0;
            
            var tbody = document.getElementById('itemTableBody');
            tbody.innerHTML = '';

            rawData.forEach(function(item) {
                // 1. 名稱過濾
                var nameMatch = !nameQuery || item.Name.toLowerCase().indexOf(nameQuery) !== -1;
                
                // 2. 屬性數值過濾
                var attrMatch = true;
                if (attrQuery) {
                    var extractedValue = parseAttributeValue(item.Attrs, attrQuery);
                    if (extractedValue !== null) {
                        attrMatch = extractedValue >= minVal;
                    } else {
                        attrMatch = false; 
                    }
                }

                if (nameMatch && attrMatch) {
                    var row = document.createElement('tr');
                    if (attrQuery) {
                        row.className = 'highlight';
                    }
                    
                    var attrDisplay = item.Attrs !== "-" ? '<span class="badge badge-attr">' + item.Attrs + '</span>' : "-";
                    
                    // 使用標準字串拼接，完美避開 PowerShell 雙引號轉譯 Bug
                    row.innerHTML = '<td><strong>' + item.Name + '</strong></td>' +
                                    '<td>' + attrDisplay + '</td>' +
                                    '<td style="color: #e53e3e; font-weight: bold;">' + item.Price + '</td>' +
                                    '<td>' + item.Qty + '</td>' +
                                    '<td>' + item.Stall + '</td>' +
                                    '<td style="color: #718096; font-size: 0.9em;">' + item.Position + '</td>' +
                                    '<td><span class="badge">' + item.Category + '</span></td>';
                    
                    tbody.appendChild(row);
                }
            });
        }

        // 初始化渲染
        filterData();
    </script>
</body>
</html>
"@

[System.IO.File]::WriteAllText($HtmlPath, $HtmlContent, [System.Text.Encoding]::UTF8)
Write-Host "HTML Dashboard updated -> $HtmlPath" -ForegroundColor Green

# 自動在瀏覽器中開啟產生的網頁
try {
    Start-Process $HtmlPath
    Write-Host "Automatically opened dashboard in your browser." -ForegroundColor Cyan
} catch {
    Write-Warning "Could not automatically open dashboard: $($_.Exception.Message)"
}

# ==================== 3. 初始化與 Demo 邏輯 ====================
if ($Init) {
    Save-State -Snapshot $snap
    Write-Host "Baseline initialized. Re-run after the site update time changes."
    exit 0
}

$prev = Load-State

if ($DemoNotify) {
    if (-not $prev) {
        Save-State -Snapshot $snap
        $prev = Load-State
    }
    $fake = [pscustomobject]@{
        Fingerprint = "DEMO|item|-|1|1|demo-stall|0, 0"
        Name        = "[DEMO] New item notify test"
        Category    = "item"
        Attrs       = "-"
        Qty         = "1"
        Price       = "1"
        Stall       = "demo-stall"
        Position    = "0, 0"
    }
    $snap.Items = @($fake) + @($snap.Items)
    $snap.Count = @($snap.Items).Count
    Write-Host "Demo mode: injecting a fake new item to trigger notify." -ForegroundColor Magenta
}

if (-not $prev) {
    Save-State -Snapshot $snap
    Write-Host "No previous state. Baseline saved. Run again later to detect new items."
    exit 0
}

# ==================== 4. 新上架比對與通知邏輯 ====================
$prevSet = [System.Collections.Generic.HashSet[string]]::new()
foreach ($it in @($prev.Items)) {
    if ($it.Fingerprint) { [void]$prevSet.Add([string]$it.Fingerprint) }
}

$newItems = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($it in @($snap.Items)) {
    $fp = [string]$it.Fingerprint
    if ($fp -and -not $prevSet.Contains($fp)) {
        $newItems.Add($it)
    }
}

$prevUpdated = [string]$prev.SiteUpdated
$nowUpdated = [string]$snap.SiteUpdated
Write-Host ("Previous site update time: {0}" -f $(if ($prevUpdated) { $prevUpdated } else { "(unknown)" }))
Write-Host ("Current  site update time: {0}" -f $(if ($nowUpdated) { $nowUpdated } else { "(unknown)" }))

if ($newItems.Count -eq 0) {
    Write-Host "No new items." -ForegroundColor Green
    if ($prevUpdated -and $nowUpdated -and ($prevUpdated -eq $nowUpdated)) {
        Write-Host "Site data timestamp unchanged -- wait for refresh, then run again." -ForegroundColor DarkYellow
    }
    Save-State -Snapshot $snap
    Write-Host "Done."
    exit 0
}

# 有新物品時的通知流程
$lines = [System.Collections.Generic.List[string]]::new()
foreach ($n in ($newItems | Select-Object -First 8)) {
    $attrDisplay = if ($n.Attrs -and $n.Attrs -ne "-") { " ($($n.Attrs))" } else { "" }
    $lines.Add("$($n.Name)$attrDisplay | $($n.Price) | $($n.Stall) @ $($n.Position)")
}
$body = "Found {0} new item(s)`n{1}" -f $newItems.Count, ($lines -join "`n")
Show-Notify -Title "Taiwan Jien - New Items" -Body $body

Write-Host "New items:"
$newItems | Select-Object Name, Attrs, Price, Stall, Position, Qty, Category | Format-Table -AutoSize

if ($DemoNotify) {
    Write-Host "Demo done. Baseline state left as previous real snapshot."
}
else {
    Save-State -Snapshot $snap
}

Write-Host "Done."
