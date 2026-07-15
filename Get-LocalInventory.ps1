<#
.SYNOPSIS
    Local Software Inventory Collector — Network Share Edition
.DESCRIPTION
    Runs locally on each target machine (e.g. via Scheduled Task).
    Collects installed software (registry) and Windows updates (WUA),
    saves a timestamped JSON snapshot and HTML report to a network share.

    No WinRM, no remote registry, no remote PowerShell — purely local
    data collection with file output to a share.

.PARAMETER SharePath
    UNC path (or drive letter) to the inventory share.
    Each machine writes into <SharePath>\<hostname>\YYYY\MM\.
.PARAMETER ExportCsv
    Export software and updates to CSV files alongside the JSON/HTML.
.PARAMETER LogPath
    Full path to the local error log file.
    Defaults to <script directory>\errors.log.
.EXAMPLE
    .\Get-LocalInventory.ps1 -SharePath "\\fileserver\inventory"
.EXAMPLE
    .\Get-LocalInventory.ps1 -SharePath "I:\inventory" -ExportCsv
#>

[CmdletBinding()]
param(
    [string]$SharePath = "\\fileserver\inventory",
    [switch]$ExportCsv,
    [string]$LogPath
)

# ============================================================
# CONFIGURATION
# ============================================================
# $SharePath is set via parameter above.  Change the default
# value if you want a different UNC path or drive letter.
# ============================================================

$scriptVersion = '1.0'
$computerName  = $env:COMPUTERNAME
$dataPath      = Join-Path $SharePath "inventory"

# ---------------------------------------------------------------
# Resolve the directory this script lives in
# ---------------------------------------------------------------
function Get-ScriptDirectory {
    if ($script:MyInvocation.MyCommand.Path) {
        Split-Path -Parent $script:MyInvocation.MyCommand.Path
    } elseif ($hostinvocation -and $hostinvocation.MyCommand.Path) {
        Split-Path -Parent $hostinvocation.MyCommand.Path
    } else {
        (Get-Location).Path
    }
}

$scriptDir = Get-ScriptDirectory

# Default log path
if (-not $LogPath) {
    $LogPath = Join-Path $scriptDir "errors.log"
}

# ---------------------------------------------------------------
# Write-Log — append a timestamped entry to the local log file
# ---------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = 'ERROR')
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $entry = "[$timestamp] [$Level] $Message"
    try {
        $entry | Out-File -FilePath $LogPath -Append -Encoding utf8
    } catch {
        # If we can't even write the log, nothing more we can do
        Write-Warning "Failed to write to log file: $_"
    }
}

# ============================================================
# HELPER FUNCTIONS — copied verbatim from Get-SoftwareInventory.ps1
# ============================================================

# ---------------------------------------------------------------
# HTML-encode helper (no external dependency)
# ---------------------------------------------------------------
function ConvertTo-HtmlEncoded {
    param([string]$Value)
    if (-not $Value) { return '' }
    $Value = $Value -replace '&', '&amp;'
    $Value = $Value -replace '<', '&lt;'
    $Value = $Value -replace '>', '&gt;'
    $Value = $Value -replace '"', '&quot;'
    $Value = $Value -replace "'", '&#39;'
    $Value
}

# ---------------------------------------------------------------
# Convert a computer name to a safe folder name for filesystem paths
# ---------------------------------------------------------------
function ConvertTo-SafeFolderName {
    param([string]$Name)
    $folder = $Name -replace '^\.$', 'localhost'
    $folder -replace '[/\\:*?"<>|]', '_'
}

# ============================================================
# DATA COLLECTION FUNCTIONS — copied verbatim from Get-SoftwareInventory.ps1
# ============================================================

# ---------------------------------------------------------------
# Get-LocalSoftware — reads installed software from registry
# Lines 169-202 of Get-SoftwareInventory.ps1
# ---------------------------------------------------------------
function Get-LocalSoftware {
    $keys = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $items = @()
    foreach ($path in $keys) {
        if (-not (Test-Path $path)) { continue }
        $arch = if ($path -match 'WOW6432Node') { '32-bit' } elseif ($path -match '^HKCU') { 'User' } else { '64-bit' }
        Get-ItemProperty $path | ForEach-Object {
            $dispName = $_.DisplayName
            if (-not $dispName) { return }
            if ($_.DisplayName -match '^\s*$') { return }
            $installDate = $_.InstallDate
            if ($installDate -and $installDate -match '^\d{8}$') {
                try { $installDate = [datetime]::ParseExact($installDate, 'yyyyMMdd', $null).ToString('yyyy-MM-dd') }
                catch { $installDate = $installDate }
            } elseif (-not $installDate) {
                $installDate = 'Unknown'
            }
            $items += [PSCustomObject]@{
                Name          = $dispName
                Version       = if ($_.DisplayVersion) { $_.DisplayVersion } else { 'Unknown' }
                Publisher     = if ($_.Publisher) { $_.Publisher } else { 'Unknown' }
                InstallDate   = $installDate
                Architecture  = $arch
                UninstallKey  = $_.PSChildName
            }
        }
    }
    Merge-SoftwareDuplicates $items
}

# ---------------------------------------------------------------
# Merge-SoftwareDuplicates — removes duplicate software entries
# Lines 210-224 of Get-SoftwareInventory.ps1
# ---------------------------------------------------------------
function Merge-SoftwareDuplicates {
    param([PSObject[]]$Items)
    if (-not $Items -or $Items.Count -eq 0) { return @() }

    $Items | Group-Object -Property { ($_.Name -replace '\s+', ' ').Trim().ToLower() } | ForEach-Object {
        $group = $_.Group
        if ($group.Count -eq 1) { $group }
        else {
            $group | Sort-Object { $_.Version -ne 'Unknown' },
                                 { $_.Architecture -eq '64-bit' },
                                 { $_.Architecture -eq '32-bit' } -Descending |
                Select-Object -First 1
        }
    } | Sort-Object Name
}

# ---------------------------------------------------------------
# Merge-UpdateDuplicates — removes duplicate update entries
# Lines 226-239 of Get-SoftwareInventory.ps1
# ---------------------------------------------------------------
function Merge-UpdateDuplicates {
    param([PSObject[]]$Items)
    if (-not $Items -or $Items.Count -eq 0) { return @() }

    $Items | Group-Object -Property { ($_.Title -replace '\s+', ' ').Trim().ToLower() } | ForEach-Object {
        $group = $_.Group
        if ($group.Count -eq 1) { $group }
        else {
            $group | Sort-Object {
                if ($_.InstallDate -eq 'Unknown') { '0000-00-00' } else { $_.InstallDate }
            } -Descending | Select-Object -First 1
        }
    } | Sort-Object Title
}

# ---------------------------------------------------------------
# Get-LocalUpdates — queries Windows Update Agent history
# Lines 355-401 of Get-SoftwareInventory.ps1
# ---------------------------------------------------------------
function Get-LocalUpdates {
    try {
        $sessionObj = New-Object -ComObject Microsoft.Update.Session
        $searcher = $sessionObj.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        if ($historyCount -eq 0) { return @() }

        $allUpdates = $searcher.QueryHistory(0, $historyCount)
        $updates = @()
        for ($i = 0; $i -lt $allUpdates.Count; $i++) {
            try {
                $entry = $allUpdates.Item($i)
                $resultCode = switch ($entry.ResultCode) {
                    0 { 'NotStarted' }
                    1 { 'InProgress' }
                    2 { 'Succeeded' }
                    3 { 'SucceededWithErrors' }
                    4 { 'Failed' }
                    5 { 'Aborted' }
                    default { 'Unknown' }
                }

                $title = if ($entry.Title) { $entry.Title } else { 'Unknown' }
                $date = if ($entry.Date) {
                    try { $entry.Date.ToString('yyyy-MM-dd') } catch { 'Unknown' }
                } else { 'Unknown' }

                $updates += [PSCustomObject]@{
                    Title       = $title
                    InstallDate = $date
                    Result      = $resultCode
                    Description = if ($entry.Description) { $entry.Description } else { '' }
                    UpdateID    = if ($entry.UpdateID) { $entry.UpdateID } else { '' }
                }
            } catch {
                # skip bad entry
            }
        }

        $updates | Where-Object { $_.Result -eq 'Succeeded' -or $_.Result -eq 'SucceededWithErrors' } |
            Sort-Object InstallDate -Descending
    } catch {
        Write-Warning "WUA query failed: $_"
        Write-Warning "  Falling back to Win32_QuickFixEngineering..."
        Get-LocalHotfixFallback
    }
}

# ---------------------------------------------------------------
# Get-LocalHotfixFallback — fallback using Get-HotFix
# Lines 408-423 of Get-SoftwareInventory.ps1
# ---------------------------------------------------------------
function Get-LocalHotfixFallback {
    try {
        Get-HotFix | ForEach-Object {
            [PSCustomObject]@{
                Title       = "$($_.HotFixID) - $($_.Description)"
                InstallDate = if ($_.InstalledOn) { $_.InstalledOn.ToString('yyyy-MM-dd') } else { 'Unknown' }
                Result      = 'Succeeded'
                Description = if ($_.Description) { $_.Description } else { '' }
                UpdateID    = if ($_.HotFixID) { $_.HotFixID } else { '' }
            }
        } | Sort-Object InstallDate -Descending
    } catch {
        Write-Warning "Hotfix fallback also failed: $_"
        @()
    }
}

# ============================================================
# SNAPSHOT — copied verbatim from Get-SoftwareInventory.ps1
# ============================================================

# ---------------------------------------------------------------
# Save-HistorySnapshot — saves inventory data as timestamped JSON
# Lines 429-488 of Get-SoftwareInventory.ps1
# ---------------------------------------------------------------
function Save-HistorySnapshot {
    param(
        [string]$Computer,
        [PSObject[]]$Software,
        [PSObject[]]$Updates,
        [string]$HistoryRoot,
        [string]$TargetYear,
        [string]$TargetMonth
    )

    if ($TargetYear -and $TargetMonth) {
        $year = $TargetYear
        $month = $TargetMonth
        $timestamp = "$TargetYear$TargetMonth-010000"
        $snapDate = Get-Date "$TargetYear-$TargetMonth-01 00:00:00"
    } else {
        $now = Get-Date
        $year = $now.ToString('yyyy')
        $month = $now.ToString('MM')
        $timestamp = $now.ToString('yyyyMMdd-HHmm')
        $snapDate = $now
    }

    $compFolder = ConvertTo-SafeFolderName $Computer

    $snapshotDir = [System.IO.Path]::Combine($HistoryRoot, $compFolder, $year, $month)
    if (-not (Test-Path $snapshotDir)) {
        New-Item -Path $snapshotDir -ItemType Directory -Force | Out-Null
    }

    $snapshotFile = Join-Path $snapshotDir "snapshot-$timestamp.json"

    $snapshot = @{
        ScriptVersion = $scriptVersion
        Computer      = $Computer
        Date          = $snapDate.ToString('yyyy-MM-dd HH:mm:ss')
        SoftwareCount = $Software.Count
        UpdateCount   = $Updates.Count
        Software      = $Software | ForEach-Object {
            @{
                Name         = $_.Name
                Version      = $_.Version
                Publisher    = $_.Publisher
                InstallDate  = $_.InstallDate
                Architecture = $_.Architecture
            }
        }
        Updates = $Updates | ForEach-Object {
            @{
                Title       = $_.Title
                InstallDate = $_.InstallDate
                Result      = $_.Result
            }
        }
    }

    $snapshot | ConvertTo-Json -Depth 4 | Out-File -FilePath $snapshotFile -Encoding utf8
    Write-Host "  Snapshot saved: $snapshotFile"
    $snapshotFile
}

# ============================================================
# CSV EXPORT — adapted from Get-SoftwareInventory.ps1
# Lines 815-843 (adjusted to write to share host folder)
# ============================================================
function New-CsvExport {
    param(
        [string]$Computer,
        [PSObject[]]$Software,
        [PSObject[]]$Updates,
        [string]$ShareRoot,
        [string]$Year = '',
        [string]$Month = ''
    )

    if (-not $Year) { $Year = (Get-Date).ToString('yyyy') }
    if (-not $Month) { $Month = (Get-Date).ToString('MM') }
    $timestamp = (Get-Date).ToString('yyyyMMdd-HHmm')

    $compFolder = ConvertTo-SafeFolderName $Computer
    $outDir = [System.IO.Path]::Combine($ShareRoot, $compFolder, $Year, $Month)
    if (-not (Test-Path $outDir)) {
        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
    }
    $swFile = Join-Path $outDir "software-$timestamp.csv"
    $Software | Select-Object Name, Version, Publisher, InstallDate, Architecture |
        Export-Csv -Path $swFile -NoTypeInformation -Encoding utf8
    Write-Host "  CSV saved: $swFile"

    $upFile = Join-Path $outDir "updates-$timestamp.csv"
    $Updates | Select-Object Title, InstallDate, Result |
        Export-Csv -Path $upFile -NoTypeInformation -Encoding utf8
    Write-Host "  CSV saved: $upFile"
}

# ============================================================
# PER-HOST HTML REPORT — adapted from New-InventoryHtmlReport
# Original: lines 587-809 of Get-SoftwareInventory.ps1
# Adapted: back-link points to share root index.html,
#          removed archive-specific links
# ============================================================
function New-LocalHtmlReport {
    param(
        [string]$Computer,
        [PSObject[]]$Software,
        [PSObject[]]$Updates,
        [string]$OutputDir,
        [string]$Year = '',
        [string]$Month = ''
    )

    $reportDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    if (-not $Year) { $Year = (Get-Date).ToString('yyyy') }
    if (-not $Month) { $Month = (Get-Date).ToString('MM') }
    $year = $Year
    $month = $Month

    $totalSw = $Software.Count
    $totalUp = $Updates.Count

    # Date anchor
    $anchor = "$year-$month"

    # Build software table rows
    $swSb = New-Object System.Text.StringBuilder
    foreach ($item in $Software) {
        $date = if ($item.InstallDate -and $item.InstallDate -ne 'Unknown') {
            $item.InstallDate
        } else { 'Unknown' }
        $null = $swSb.Append(@"
<tr><td>$(ConvertTo-HtmlEncoded $item.Name)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Version)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Publisher)</td>
    <td>$(ConvertTo-HtmlEncoded $date)</td></tr>
"@)
    }
    $swRows = $swSb.ToString()

    # Build update table rows
    $upSb = New-Object System.Text.StringBuilder
    foreach ($item in $Updates) {
        $date = if ($item.InstallDate -and $item.InstallDate -ne 'Unknown') {
            $item.InstallDate
        } else { 'Unknown' }
        $title = $item.Title
        $null = $upSb.Append(@"
<tr><td>$(ConvertTo-HtmlEncoded $title)</td>
    <td>$(ConvertTo-HtmlEncoded $date)</td></tr>
"@)
    }
    $upRows = $upSb.ToString()

    # Back-link: from <host>/YYYY/MM/report.html to share root index.html
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Software Inventory - $(ConvertTo-HtmlEncoded $Computer) - $anchor</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
  h1, h2, h3 { color: #1a3a5c; }
  .summary { background: #fff; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 20px; }
  .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px,1fr)); gap: 10px; }
  .summary-item { background: #e8f0fe; padding: 10px; border-radius: 4px; text-align: center; }
  .summary-item .number { font-size: 24px; font-weight: bold; color: #1a3a5c; }
  .summary-item .label { font-size: 13px; color: #666; }
  .meta { font-size: 13px; color: #888; margin-top: 10px; }
  table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 25px; }
  th { background: #1a3a5c; color: #fff; padding: 10px 12px; text-align: left; font-weight: 600; cursor: pointer; }
  th:hover { background: #2a5a8c; }
  td { padding: 8px 12px; border-bottom: 1px solid #e0e0e0; word-break: break-word; }
  tr:hover td { background: #f0f5ff; }
  .section-title { margin: 25px 0 10px 0; }
  .badge-new { display: inline-block; background: #2e7d32; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
  .badge-update { display: inline-block; background: #1565c0; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
  .search-box { margin-bottom: 10px; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; width: 300px; max-width: 100%; box-sizing: border-box; }
  .search-box:focus { outline: none; border-color: #1a3a5c; box-shadow: 0 0 4px rgba(26,58,92,.3); }
  a.back-link { color: #1a3a5c; text-decoration: none; }
  a.back-link:hover { text-decoration: underline; }
  .theme-toggle { float: right; background: none; border: 1px solid #1a3a5c; color: #1a3a5c; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 13px; }
  .theme-toggle:hover { background: #1a3a5c; color: #fff; }
  @media (prefers-color-scheme: dark) {
    html.theme-auto body { background: #1a1a2e; color: #e0e0e0; }
    html.theme-auto body h1, html.theme-auto body h2, html.theme-auto body h3 { color: #80b0e0; }
    html.theme-auto body .summary { background: #16213e; }
    html.theme-auto body .summary-item { background: #0f3460; }
    html.theme-auto body .summary-item .number { color: #80b0e0; }
    html.theme-auto body .summary-item .label { color: #a0c0e0; }
    html.theme-auto body .meta { color: #888; }
    html.theme-auto body table { background: #16213e; }
    html.theme-auto body th { background: #0f3460; }
    html.theme-auto body th:hover { background: #1a4a7a; }
    html.theme-auto body td { border-bottom: 1px solid #2a3a5e; }
    html.theme-auto body tr:hover td { background: #1a2a4e; }
    html.theme-auto body .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
    html.theme-auto body .search-box:focus { border-color: #80b0e0; }
  }
  html.dark body { background: #1a1a2e; color: #e0e0e0; }
  html.dark body h1, html.dark body h2, html.dark body h3 { color: #80b0e0; }
  html.dark body .summary { background: #16213e; }
  html.dark body .summary-item { background: #0f3460; }
  html.dark body .summary-item .number { color: #80b0e0; }
  html.dark body .summary-item .label { color: #a0c0e0; }
  html.dark body .meta { color: #888; }
  html.dark body table { background: #16213e; }
  html.dark body th { background: #0f3460; }
  html.dark body th:hover { background: #1a4a7a; }
  html.dark body td { border-bottom: 1px solid #2a3a5e; }
  html.dark body tr:hover td { background: #1a2a4e; }
  html.dark body .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
  html.dark body .search-box:focus { border-color: #80b0e0; }
</style>
<script>
function toggleTheme() {
  var theme = document.documentElement.classList.contains('dark') ? 'light' : 'dark';
  document.documentElement.classList.toggle('dark');
  document.documentElement.classList.remove('theme-auto');
  localStorage.setItem('theme', theme);
  var links = document.querySelectorAll('a');
  for (var i = 0; i < links.length; i++) {
    var href = links[i].getAttribute('href');
    if (!href || href.indexOf('://') >= 0 || href.indexOf('#') >= 0) continue;
    href = href.replace(/[?&]theme=\w+/g, '');
    href += (href.indexOf('?') >= 0 ? '&' : '?') + 'theme=' + theme;
    links[i].setAttribute('href', href);
  }
}
(function() {
  try {
    var m = window.location.search.match(/[?&]theme=(\w+)/);
    if (m && m[1] === 'dark') { document.documentElement.classList.add('dark'); return; }
    if (m && m[1] === 'light') { document.documentElement.classList.remove('dark', 'theme-auto'); return; }
    var saved = localStorage.getItem('theme');
    if (saved === 'dark') { document.documentElement.classList.add('dark'); return; }
    if (saved === 'light') { document.documentElement.classList.remove('dark', 'theme-auto'); return; }
    if (window.matchMedia('(prefers-color-scheme: dark)').matches) document.documentElement.classList.add('theme-auto');
  } catch(e) {
    if (window.matchMedia('(prefers-color-scheme: dark)').matches) document.documentElement.classList.add('dark');
  }
})();
document.addEventListener('click',function(e){
  var el=e.target;
  while(el&&el.tagName!=='A')el=el.parentNode;
  if(!el)return;
  var href=el.getAttribute('href');
  if(!href||href.indexOf('://')>=0||href.indexOf('#')>=0||href.indexOf('?theme=')>=0)return;
  var theme=document.documentElement.classList.contains('dark')?'dark':'light';
  el.href=href+(href.indexOf('?')>=0?'&':'?')+'theme='+theme;
});
function filterTable(inputId, tableId) {
  var input = document.getElementById(inputId);
  var filter = input.value.toLowerCase();
  var table = document.getElementById(tableId);
  var tbody = table.querySelector('tbody');
  if (!tbody) return;
  var rows = tbody.querySelectorAll('tr');
  for (var i = 0; i < rows.length; i++) {
    var text = rows[i].textContent.toLowerCase();
    rows[i].style.display = text.indexOf(filter) > -1 ? '' : 'none';
  }
}
function sortTable(tableId, col) {
  var table = document.getElementById(tableId);
  var tbody = table.querySelector('tbody');
  if (!tbody) return;
  var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
  var dir = table.getAttribute('data-sort-dir-' + col) === 'asc' ? 'desc' : 'asc';
  table.setAttribute('data-sort-dir-' + col, dir);
  var multiplier = dir === 'asc' ? 1 : -1;
  rows.sort(function(a, b) {
    var aText = a.children[col].textContent.trim();
    var bText = b.children[col].textContent.trim();
    var aDate = Date.parse(aText);
    var bDate = Date.parse(bText);
    if (!isNaN(aDate) && !isNaN(bDate)) return (aDate - bDate) * multiplier;
    return aText.localeCompare(bText, undefined, { numeric: true }) * multiplier;
  });
  rows.forEach(function(row) { tbody.appendChild(row); });
}
</script>
</head>
<body>
<button class="theme-toggle" onclick="toggleTheme()">&#9681; Theme</button>
<a class="back-link" href="../../../../index.html">&larr; Back to inventory</a>
<h1>Software Inventory Report</h1>
<div class="summary">
  <div class="summary-grid">
    <div class="summary-item"><div class="number">$totalSw</div><div class="label">3rd Party Software</div></div>
    <div class="summary-item"><div class="number">$totalUp</div><div class="label">Windows Patches</div></div>
  </div>
  <div class="meta">Computer: <strong>$(ConvertTo-HtmlEncoded $Computer)</strong> &nbsp;|&nbsp; Generated: $reportDate &nbsp;|&nbsp; Snapshot: $anchor</div>
</div>
"@

    # All Software
    $html += @"
<h2 class="section-title">3rd Party Software <span class="badge-new">$totalSw</span></h2>
<input type="text" id="allSw-filter" class="search-box" placeholder="Filter software..." onkeyup="filterTable('allSw-filter','allSw-table')">
<table id="allSw-table"><thead><tr><th onclick="sortTable('allSw-table',0)">Name</th><th onclick="sortTable('allSw-table',1)">Version</th><th onclick="sortTable('allSw-table',2)">Publisher</th><th onclick="sortTable('allSw-table',3)">Install Date</th></tr></thead><tbody>$swRows</tbody></table>
"@

    # All Updates
    $html += @"
<h2 class="section-title">Windows Patches <span class="badge-update">$totalUp</span></h2>
<input type="text" id="allUp-filter" class="search-box" placeholder="Filter updates..." onkeyup="filterTable('allUp-filter','allUp-table')">
<table id="allUp-table"><thead><tr><th onclick="sortTable('allUp-table',0)">Title</th><th onclick="sortTable('allUp-table',1)">Install Date</th></tr></thead><tbody>$upRows</tbody></table>
</body></html>
"@

    # Ensure output directory exists
    $compFolder = ConvertTo-SafeFolderName $Computer
    $outDir = [System.IO.Path]::Combine($OutputDir, $compFolder, $year, $month)
    if (-not (Test-Path $outDir)) {
        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
    }
    $reportFile = Join-Path $outDir "report.html"
    $html | Out-File -FilePath $reportFile -Encoding utf8
    Write-Host "  Report saved: $reportFile"
    $reportFile
}

# ============================================================
# MAIN EXECUTION
# ============================================================
Write-Host "Local Software Inventory Collector v$scriptVersion"
Write-Host "Computer: $computerName"
Write-Host "Share:    $SharePath"
Write-Host ""

# ---------------------------------------------------------------
# Step 1 — Validate share is accessible and writable
# ---------------------------------------------------------------
Write-Host "Validating share accessibility..."
$testFile = Join-Path $SharePath "_write_test_$($computerName).tmp"

try {
    # First check the share path exists
    if (-not (Test-Path $SharePath)) {
        throw "Share path not found: $SharePath"
    }

    # Try writing a temp file to prove write access
    "write-test" | Out-File -FilePath $testFile -Encoding utf8 -ErrorAction Stop
    Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
    Write-Host "  Share is accessible and writable."
} catch {
    $msg = "Share validation failed for $computerName : $SharePath — $_"
    Write-Warning $msg
    Write-Log $msg
    exit 1
}

# ---------------------------------------------------------------
# Step 2 — Collect installed software (registry)
# ---------------------------------------------------------------
Write-Host "Collecting software..."
$software = @()
try {
    $software = Get-LocalSoftware
    Write-Host "  Found $($software.Count) software entries."
} catch {
    $msg = "Software collection failed for $computerName : $_"
    Write-Warning $msg
    Write-Log $msg
}

# ---------------------------------------------------------------
# Step 3 — Collect installed updates (WUA COM + fallback)
# ---------------------------------------------------------------
Write-Host "Collecting updates..."
$updates = @()
try {
    $updates = Get-LocalUpdates
    if (-not $updates -or $updates.Count -eq 0) {
        Write-Warning "  No updates from WUA, trying hotfix fallback..."
        $updates = Get-LocalHotfixFallback
    }
    $updates = Merge-UpdateDuplicates $updates
    Write-Host "  Found $($updates.Count) updates."
} catch {
    $msg = "Update collection failed for $computerName : $_"
    Write-Warning $msg
    Write-Log $msg
}

# ---------------------------------------------------------------
# Step 4 — Save JSON snapshot to share
# ---------------------------------------------------------------
Write-Host "Saving snapshot..."
try {
    Save-HistorySnapshot -Computer $computerName -Software $software -Updates $updates -HistoryRoot $dataPath
} catch {
    $msg = "Snapshot save failed for $computerName : $_"
    Write-Warning $msg
    Write-Log $msg
}

# ---------------------------------------------------------------
# Step 5 — Generate per-host HTML report on share
# ---------------------------------------------------------------
Write-Host "Generating HTML report..."
try {
    New-LocalHtmlReport -Computer $computerName -Software $software -Updates $updates -OutputDir $dataPath
} catch {
    $msg = "HTML report generation failed for $computerName : $_"
    Write-Warning $msg
    Write-Log $msg
}

# ---------------------------------------------------------------
# Step 6 — Optional CSV export
# ---------------------------------------------------------------
if ($ExportCsv) {
    Write-Host "Exporting CSV..."
    try {
        New-CsvExport -Computer $computerName -Software $software -Updates $updates -ShareRoot $dataPath
    } catch {
        $msg = "CSV export failed for $computerName : $_"
        Write-Warning $msg
        Write-Log $msg
    }
}

Write-Host ""
Write-Host "Done. Inventory for $computerName written to $SharePath"
