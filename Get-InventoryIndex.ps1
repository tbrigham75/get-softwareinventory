<#
.SYNOPSIS
    Central Inventory Index Generator — Network Share Edition
.DESCRIPTION
    Scans the inventory share for host folders, reads each host's
    latest JSON snapshot, and generates a central index.html with
    links to each host's report.

    No WinRM, no remote calls — reads/writes only files on the share.

.PARAMETER SharePath
    UNC path (or drive letter) to the inventory share.
    Scans for host subfolders and generates index.html at the root.
.PARAMETER LogPath
    Full path to the local error log file.
    Defaults to <script directory>\errors.log.
.EXAMPLE
    .\Get-InventoryIndex.ps1 -SharePath "\\fileserver\inventory"
.EXAMPLE
    .\Get-InventoryIndex.ps1 -SharePath "I:\inventory" -LogPath "C:\Logs\index-errors.log"
#>

[CmdletBinding()]
param(
    [string]$SharePath = "\\fileserver\inventory",
    [string]$LogPath
)

# ============================================================
# CONFIGURATION
# ============================================================
# $SharePath is set via parameter above.  Change the default
# value if you want a different UNC path or drive letter.
# ============================================================

$scriptVersion = '1.0'

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
        Write-Warning "Failed to write to log file: $_"
    }
}

# ============================================================
# HELPER FUNCTIONS — copied verbatim from Get-SoftwareInventory.ps1
# ============================================================

# ---------------------------------------------------------------
# HTML-encode helper
# Lines 62-71 of Get-SoftwareInventory.ps1
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
# Convert a computer name to a safe folder name
# Lines 74-78 of Get-SoftwareInventory.ps1
# ---------------------------------------------------------------
function ConvertTo-SafeFolderName {
    param([string]$Name)
    $folder = $Name -replace '^\.$', 'localhost'
    $folder -replace '[/\\:*?"<>|]', '_'
}

# ============================================================
# SNAPSHOT LOADER — copied verbatim from Get-SoftwareInventory.ps1
# ============================================================

# ---------------------------------------------------------------
# Load-HistorySnapshot — loads the most recent JSON snapshot
# Lines 495-517 of Get-SoftwareInventory.ps1
# ---------------------------------------------------------------
function Load-HistorySnapshot {
    param(
        [string]$Computer,
        [string]$HistoryRoot
    )

    $compFolder = ConvertTo-SafeFolderName $Computer
    $historyDir = Join-Path $HistoryRoot $compFolder
    if (-not (Test-Path $historyDir)) { return $null }

    $snapshotFiles = Get-ChildItem -Path $historyDir -Recurse -Filter 'snapshot-*.json' |
        Sort-Object LastWriteTime -Descending

    if ($snapshotFiles.Count -eq 0) { return $null }

    try {
        $data = Get-Content -Path $snapshotFiles[0].FullName -Raw | ConvertFrom-Json
        $data
    } catch {
        Write-Warning "Could not load previous snapshot: $_"
        $null
    }
}

# ============================================================
# MAIN EXECUTION
# ============================================================
Write-Host "Inventory Index Generator v$scriptVersion"
Write-Host "Share: $SharePath"
Write-Host ""

# ---------------------------------------------------------------
# Step 1 — Validate share path
# ---------------------------------------------------------------
if (-not (Test-Path $SharePath)) {
    $msg = "Share path not found: $SharePath"
    Write-Warning $msg
    Write-Log $msg
    exit 1
}

# ---------------------------------------------------------------
# Step 2 — Scan for host folders
# ---------------------------------------------------------------
Write-Host "Scanning for host folders..."
$hostFolders = Get-ChildItem -Path $SharePath -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -ne '_logs' -and $_.Name -notmatch '^\.' }

Write-Host "  Found $($hostFolders.Count) host folder(s)."

if ($hostFolders.Count -eq 0) {
    Write-Warning "No host folders found in $SharePath — nothing to index."
    exit 0
}

# ---------------------------------------------------------------
# Step 3 — Read latest snapshot from each host
# ---------------------------------------------------------------
Write-Host "Reading host snapshots..."
$hostEntries = @()
$totalSoftware = 0
$totalUpdates = 0

foreach ($folder in $hostFolders) {
    $hostName = $folder.Name
    try {
        $snapshot = Load-HistorySnapshot -Computer $hostName -HistoryRoot $SharePath
        if ($snapshot) {
            $swCount = if ($snapshot.SoftwareCount) { $snapshot.SoftwareCount } else { 0 }
            $upCount = if ($snapshot.UpdateCount) { $snapshot.UpdateCount } else { 0 }
            $totalSoftware += $swCount
            $totalUpdates += $upCount

            # Find the latest report.html for this host
            $reportFiles = Get-ChildItem -Path $folder.FullName -Recurse -Filter 'report.html' -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending
            $reportRelPath = $null
            if ($reportFiles.Count -gt 0) {
                # Relative path from share root to report
                $reportRelPath = $reportFiles[0].FullName.Substring($SharePath.Length + 1)
            }

            $hostEntries += [PSCustomObject]@{
                HostName      = $hostName
                SoftwareCount = $swCount
                UpdateCount   = $upCount
                LastScan      = $snapshot.Date
                ReportPath    = $reportRelPath
            }
            Write-Host "  $hostName : $swCount software, $upCount updates (last: $($snapshot.Date))"
        } else {
            Write-Warning "  $hostName : no snapshot found — skipping."
        }
    } catch {
        Write-Warning "  $hostName : failed to read snapshot — $_"
        Write-Log "Failed to read snapshot for $hostName : $_"
    }
}

Write-Host ""
Write-Host "  Total: $($hostEntries.Count) host(s), $totalSoftware software entries, $totalUpdates update entries"

# ---------------------------------------------------------------
# Step 4 — Build index.html
# ---------------------------------------------------------------
Write-Host "Generating index.html..."

$reportDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
$lastScanDate = ($hostEntries | Sort-Object LastScan -Descending | Select-Object -First 1).LastScan
if (-not $lastScanDate) { $lastScanDate = 'N/A' }

# Build table rows
$hostSb = New-Object System.Text.StringBuilder
foreach ($entry in ($hostEntries | Sort-Object HostName)) {
    $hostLink = if ($entry.ReportPath) {
        "<a href=""$($entry.ReportPath)"">$(ConvertTo-HtmlEncoded $entry.HostName)</a>"
    } else {
        ConvertTo-HtmlEncoded $entry.HostName
    }
    $scanDate = if ($entry.LastScan) { $entry.LastScan } else { 'N/A' }
    $null = $hostSb.Append(@"
<tr>
    <td>$hostLink</td>
    <td>$($entry.SoftwareCount)</td>
    <td>$($entry.UpdateCount)</td>
    <td>$(ConvertTo-HtmlEncoded $scanDate)</td>
</tr>
"@)
}
$hostRows = $hostSb.ToString()

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Software Inventory - Network Index</title>
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
  .badge { display: inline-block; background: #1a3a5c; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
  .search-box { margin-bottom: 10px; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; width: 300px; max-width: 100%; box-sizing: border-box; }
  .search-box:focus { outline: none; border-color: #1a3a5c; box-shadow: 0 0 4px rgba(26,58,92,.3); }
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
<h1>Software Inventory &mdash; Network Index</h1>
<div class="summary">
  <div class="summary-grid">
    <div class="summary-item"><div class="number">$($hostEntries.Count)</div><div class="label">Hosts</div></div>
    <div class="summary-item"><div class="number">$totalSoftware</div><div class="label">Total Software Entries</div></div>
    <div class="summary-item"><div class="number">$totalUpdates</div><div class="label">Total Update Entries</div></div>
  </div>
  <div class="meta">Generated: $reportDate &nbsp;|&nbsp; Last scan: $(ConvertTo-HtmlEncoded $lastScanDate)</div>
</div>

<h2 class="section-title">Discovered Hosts <span class="badge">$($hostEntries.Count)</span></h2>
<input type="text" id="host-filter" class="search-box" placeholder="Filter hosts..." onkeyup="filterTable('host-filter','host-table')">
<table id="host-table"><thead><tr>
    <th onclick="sortTable('host-table',0)">Hostname</th>
    <th onclick="sortTable('host-table',1)">Software</th>
    <th onclick="sortTable('host-table',2)">Updates</th>
    <th onclick="sortTable('host-table',3)">Last Scan</th>
</tr></thead><tbody>
$hostRows
</tbody></table>
</body></html>
"@

# Write index.html to share root
$indexPath = Join-Path $SharePath "index.html"
try {
    $html | Out-File -FilePath $indexPath -Encoding utf8
    Write-Host "  Index saved: $indexPath"
} catch {
    $msg = "Failed to write index.html: $_"
    Write-Warning $msg
    Write-Log $msg
    exit 1
}

Write-Host ""
Write-Host "Done. Index generated with $($hostEntries.Count) host(s)."
