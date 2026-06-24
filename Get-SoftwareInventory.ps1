<#
.SYNOPSIS
    Software Inventory and Update Reporting Script
.DESCRIPTION
    Collects installed software (from registry) and installed updates (from WUA),
    saves persistent monthly JSON snapshots, and generates browsable HTML reports
    with year/month navigation, install dates, and change tracking.
.PARAMETER ComputerName
    One or more computers to inventory. Overrides hostnames.txt if provided.
.PARAMETER OutputPath
    Root folder for HTML reports. Defaults to .\Output.
.PARAMETER HistoryPath
    Root folder for historical snapshots. Defaults to .\History.
.PARAMETER PassThru
    Return inventory objects to the pipeline instead of (or in addition to) saving files.
.PARAMETER ExportCsv
    Export software and updates to CSV files alongside HTML reports.
.PARAMETER ThrottleLimit
    Maximum number of concurrent remote host collections (default: 5, max: 50).
.EXAMPLE
    .\Get-SoftwareInventory.ps1
.EXAMPLE
    .\Get-SoftwareInventory.ps1 -ComputerName SRV01, SRV02
.EXAMPLE
    .\Get-SoftwareInventory.ps1 -ComputerName localhost -OutputPath C:\Reports
.EXAMPLE
    .\Get-SoftwareInventory.ps1 -ComputerName SRV01 -ExportCsv
#>

[CmdletBinding()]
param(
    [string[]]$ComputerName,
    [string]$OutputPath,
    [string]$HistoryPath,
    [switch]$PassThru,
    [switch]$ExportCsv,
    [ValidateRange(1, 50)]
    [int]$ThrottleLimit = 5
)

# ---------------------------------------------------------------
# Self-contained helper: resolve the directory this script lives in
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
$scriptVersion = '7.0'

# Retry constants for remote connections
$script:RetryCount = 2
$script:RetryDelayMs = 2000

# HTML-encode helper (no external dependency)
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

# Default paths (outside git repo — no system info leaked)
$webRoot = "C:\Utils\Web\get-softwareinventory"
if (-not $OutputPath)  { $OutputPath  = Join-Path $webRoot "Output" }
if (-not $HistoryPath) { $HistoryPath = Join-Path $webRoot "History" }

# Resolve computers list
if (-not $ComputerName -or $ComputerName.Count -eq 0) {
    $hostsFile = Join-Path $scriptDir "hostnames.txt"
    if (Test-Path $hostsFile) {
        $ComputerName = Get-Content $hostsFile | Where-Object { $_.Trim() -ne '' -and $_ -notmatch '^\s*#' }
    } else {
        $ComputerName = @($env:COMPUTERNAME)
    }
}

# ---------------------------------------------------------------
# Helper: is the target the local machine?
# ---------------------------------------------------------------
function Test-IsLocalComputer {
    param([string]$Computer)
    $local = $env:COMPUTERNAME
    $Computer -eq $local -or $Computer -eq '127.0.0.1' -or $Computer -eq 'localhost' -or $Computer -eq '.'
}

# ---------------------------------------------------------------
# Test-ComputerConnectivity
# Quickly checks whether a remote computer is reachable via WinRM.
# Returns $true if reachable, $false otherwise.
# ---------------------------------------------------------------
function Test-ComputerConnectivity {
    param([string]$Computer)
    if (Test-IsLocalComputer $Computer) { return $true }
    try {
        $null = Test-WSMan -ComputerName $Computer -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# ---------------------------------------------------------------
# Get-SoftwareInventory
# Reads installed software from registry Uninstall keys.
# Returns a list of PSObjects with Name, Version, Publisher, InstallDate, Architecture.
# ---------------------------------------------------------------
function Get-SoftwareInventory {
    [CmdletBinding()]
    param([string]$Computer)

    $isLocal = Test-IsLocalComputer $Computer

    if ($isLocal) {
        Get-LocalSoftware
    } else {
        $result = $null
        for ($attempt = 1; $attempt -le $script:RetryCount; $attempt++) {
            try {
                $session = New-PSSession -ComputerName $Computer -ErrorAction Stop
                $result = Invoke-Command -Session $session -ScriptBlock ${function:Get-LocalSoftware} -ErrorAction Stop
                Remove-PSSession $session
                break
            } catch {
                if ($attempt -lt $script:RetryCount) {
                    Write-Warning "  WinRM attempt $attempt failed for $Computer, retrying in $($script:RetryDelayMs)ms..."
                    Start-Sleep -Milliseconds $script:RetryDelayMs
                } else {
                    Write-Warning "Failed to connect to $Computer via WinRM after $script:RetryCount attempts: $_"
                    Write-Warning "Attempting remote registry fallback..."
                    $result = Get-RemoteRegistrySoftware $Computer
                }
            }
        }
        $result
    }
}

# ---------------------------------------------------------------
# Get-LocalSoftware (runs locally or inside Invoke-Command)
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
# Merge-SoftwareDuplicates
# Removes duplicate software entries that share the same name
# (e.g. KB patches / VC++ runtimes registered in both 32- and 64-bit
#  uninstall keys).  Keeps the entry with the richest data.
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
# Get-RemoteRegistrySoftware - fallback using .NET remote registry
# ---------------------------------------------------------------
function Get-RemoteRegistrySoftware {
    param([string]$Computer)

    try {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
    } catch {
        Write-Warning "Remote registry not available on $Computer : $_"
        return @()
    }

    $subKeys = @(
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
        'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $items = @()
    foreach ($path in $subKeys) {
        $arch = if ($path -match 'WOW6432Node') { '32-bit' } else { '64-bit' }
        try {
            $key = $reg.OpenSubKey($path)
            if (-not $key) { continue }
            $names = $key.GetSubKeyNames()
            foreach ($name in $names) {
                try {
                    $sub = $key.OpenSubKey($name)
                    $dispName = $sub.GetValue('DisplayName')
                    if (-not $dispName) { continue }
                    $installDate = $sub.GetValue('InstallDate')
                    if ($installDate -and $installDate -match '^\d{8}$') {
                        try { $installDate = [datetime]::ParseExact($installDate, 'yyyyMMdd', $null).ToString('yyyy-MM-dd') }
                        catch { }
                    } elseif (-not $installDate) {
                        $installDate = 'Unknown'
                    }
                    $items += [PSCustomObject]@{
                        Name          = $dispName
                        Version       = if ($v = $sub.GetValue('DisplayVersion')) { $v } else { 'Unknown' }
                        Publisher     = if ($p = $sub.GetValue('Publisher')) { $p } else { 'Unknown' }
                        InstallDate   = $installDate
                        Architecture  = $arch
                        UninstallKey  = $name
                    }
                } catch {
                    # skip bad key
                } finally {
                    if ($sub) { $sub.Close() }
                }
            }
        } catch {
            Write-Warning "Could not read $path on $Computer"
        } finally {
            if ($key) { $key.Close() }
        }
    }
    $reg.Close()
    Merge-SoftwareDuplicates $items
}

# ---------------------------------------------------------------
# Get-InstalledUpdates
# Queries Windows Update Agent history.
# Returns a list of PSObjects with Title, InstallDate, Result, Description.
# ---------------------------------------------------------------
function Get-InstalledUpdates {
    [CmdletBinding()]
    param([string]$Computer)

    $isLocal = Test-IsLocalComputer $Computer
    $result = @()

    if ($isLocal) {
        $result = Get-LocalUpdates
    } else {
        for ($attempt = 1; $attempt -le $script:RetryCount; $attempt++) {
            try {
                $session = New-PSSession -ComputerName $Computer -ErrorAction Stop
                $result = Invoke-Command -Session $session -ScriptBlock ${function:Get-LocalUpdates} -ErrorAction Stop
                Remove-PSSession $session
                break
            } catch {
                if ($attempt -lt $script:RetryCount) {
                    Write-Warning "  WinRM attempt $attempt failed for $Computer (updates), retrying..."
                    Start-Sleep -Milliseconds $script:RetryDelayMs
                } else {
                    Write-Warning "Failed to get updates from $Computer via WinRM after $script:RetryCount attempts: $_"
                }
            }
        }
    }

    if (-not $result -or $result.Count -eq 0) {
        Write-Warning "  Trying Win32_QuickFixEngineering fallback for $Computer..."
        if ($isLocal) {
            $result = Get-LocalHotfixFallback
        } else {
            try {
                $session = New-PSSession -ComputerName $Computer -ErrorAction Stop
                $result = Invoke-Command -Session $session -ScriptBlock ${function:Get-LocalHotfixFallback} -ErrorAction Stop
                Remove-PSSession $session
            } catch {
                Write-Warning "  Hotfix fallback also failed for $Computer : $_"
            }
        }
    }

    Merge-UpdateDuplicates $result
}

# ---------------------------------------------------------------
# Get-LocalUpdates (runs locally or inside Invoke-Command)
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
# Get-LocalHotfixFallback
# Queries Win32_QuickFixEngineering (Get-HotFix) as a fallback
# when the WUA COM API is unavailable.
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

# ---------------------------------------------------------------
# Save-HistorySnapshot
# Saves inventory data as a timestamped JSON file under HistoryPath.
# ---------------------------------------------------------------
function Save-HistorySnapshot {
    param(
        [string]$Computer,
        [PSObject[]]$Software,
        [PSObject[]]$Updates,
        [string]$HistoryRoot
    )

    $now = Get-Date
    $year = $now.ToString('yyyy')
    $month = $now.ToString('MM')
    $timestamp = $now.ToString('yyyyMMdd-HHmm')

    $compFolder = $Computer -replace '^\.$', 'localhost'
    $compFolder = $compFolder -replace '[/\\:*?"<>|]', '_'

    $snapshotDir = [System.IO.Path]::Combine($HistoryRoot, $compFolder, $year, $month)
    if (-not (Test-Path $snapshotDir)) {
        New-Item -Path $snapshotDir -ItemType Directory -Force | Out-Null
    }

    $snapshotFile = Join-Path $snapshotDir "snapshot-$timestamp.json"

    $snapshot = @{
        ScriptVersion = $scriptVersion
        Computer      = $Computer
        Date          = $now.ToString('yyyy-MM-dd HH:mm:ss')
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

# ---------------------------------------------------------------
# Load-HistorySnapshot
# Loads the most recent JSON snapshot for the given computer.
# Returns $null if no snapshot exists.
# ---------------------------------------------------------------
function Load-HistorySnapshot {
    param(
        [string]$Computer,
        [string]$HistoryRoot
    )

    $compFolder = $Computer -replace '[/\\:*?"<>|]', '_'
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

# ---------------------------------------------------------------
# Compare-Snapshots
# Compares current and previous snapshots, returning
# NewSoftware, RemovedSoftware, NewUpdates.
# ---------------------------------------------------------------
function Compare-Snapshots {
    param(
        [PSObject[]]$CurrentSoftware,
        [PSObject[]]$CurrentUpdates,
        [PSObject]$PreviousSnapshot
    )

    if (-not $PreviousSnapshot) {
        return @{
            NewSoftware    = @()
            RemovedSoftware = @()
            NewUpdates     = @()
        }
    }

    # Normalize names for comparison
    $prevSw = $PreviousSnapshot.Software | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Normalized  = ($_.Name -replace '\s+', ' ').Trim().ToLower()
            Version     = $_.Version
        }
    }

    $currSw = $CurrentSoftware | ForEach-Object {
        [PSCustomObject]@{
            Name        = $_.Name
            Normalized  = ($_.Name -replace '\s+', ' ').Trim().ToLower()
            Version     = $_.Version
            Publisher   = $_.Publisher
            InstallDate = $_.InstallDate
        }
    }

    $prevNorm = $prevSw | ForEach-Object { $_.Normalized }
    $currNorm = $currSw | ForEach-Object { $_.Normalized }

    $newSw = $currSw | Where-Object { $_.Normalized -notin $prevNorm }
    $removedSw = $prevSw | Where-Object { $_.Normalized -notin $currNorm }

    # Updates comparison
    $prevUpTitles = $PreviousSnapshot.Updates | ForEach-Object {
        ($_.Title -replace '\s+', ' ').Trim()
    }
    $currUpTitles = $CurrentUpdates | ForEach-Object {
        ($_.Title -replace '\s+', ' ').Trim()
    }

    $newUp = $CurrentUpdates | Where-Object {
        ($_.Title -replace '\s+', ' ').Trim() -notin $prevUpTitles
    }

    @{
        NewSoftware     = $newSw | Sort-Object Name
        RemovedSoftware = $removedSw | Sort-Object Name
        NewUpdates      = $newUp | Sort-Object Title
    }
}

# ---------------------------------------------------------------
# New-InventoryHtmlReport
# Generates a full HTML report page for a single computer run.
# ---------------------------------------------------------------
function New-InventoryHtmlReport {
    param(
        [string]$Computer,
        [PSObject[]]$Software,
        [PSObject[]]$Updates,
        [hashtable]$Comparison,
        [PSObject]$PreviousSnapshot,
        [string]$OutputDir
    )

    $now = Get-Date
    $year = $now.ToString('yyyy')
    $month = $now.ToString('MM')
    $reportDate = $now.ToString('yyyy-MM-dd HH:mm:ss')

    $newSwCount = if ($Comparison) { $Comparison.NewSoftware.Count } else { 0 }
    $remSwCount = if ($Comparison) { $Comparison.RemovedSoftware.Count } else { 0 }
    $newUpCount = if ($Comparison) { $Comparison.NewUpdates.Count } else { 0 }
    $totalSw = $Software.Count
    $totalUp = $Updates.Count

    # Date anchor
    $anchor = "$year-$month"

    # Build software table rows
    $swRows = ''
    foreach ($item in $Software) {
        $date = if ($item.InstallDate -and $item.InstallDate -ne 'Unknown') {
            $item.InstallDate
        } else { 'Unknown' }
        $swRows += @"
<tr><td>$(ConvertTo-HtmlEncoded $item.Name)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Version)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Publisher)</td>
    <td>$(ConvertTo-HtmlEncoded $date)</td></tr>
"@
    }

    # Build update table rows
    $upRows = ''
    foreach ($item in $Updates) {
        $date = if ($item.InstallDate -and $item.InstallDate -ne 'Unknown') {
            $item.InstallDate
        } else { 'Unknown' }
        $title = $item.Title
        $upRows += @"
<tr><td>$(ConvertTo-HtmlEncoded $title)</td>
    <td>$(ConvertTo-HtmlEncoded $date)</td></tr>
"@
    }

    # New software rows
    $newSwRows = ''
    if ($newSwCount -gt 0) {
        foreach ($item in $Comparison.NewSoftware) {
            $date = if ($item.InstallDate -and $item.InstallDate -ne 'Unknown') {
                $item.InstallDate
            } else { 'Unknown' }
            $newSwRows += @"
<tr><td>$(ConvertTo-HtmlEncoded $item.Name)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Version)</td>
    <td>$(ConvertTo-HtmlEncoded $date)</td></tr>
"@
        }
    }

    # Removed software rows
    $remSwRows = ''
    if ($remSwCount -gt 0) {
        foreach ($item in $Comparison.RemovedSoftware) {
            $remSwRows += @"
<tr><td>$(ConvertTo-HtmlEncoded $item.Name)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Version)</td></tr>
"@
        }
    }

    # New update rows
    $newUpRows = ''
    if ($newUpCount -gt 0) {
        foreach ($item in $Comparison.NewUpdates) {
            $date = if ($item.InstallDate -and $item.InstallDate -ne 'Unknown') {
                $item.InstallDate
            } else { 'Unknown' }
            $title = $item.Title
            $newUpRows += @"
<tr><td>$(ConvertTo-HtmlEncoded $title)</td>
    <td>$(ConvertTo-HtmlEncoded $date)</td></tr>
"@
        }
    }

    $prevDate = if ($PreviousSnapshot) { $PreviousSnapshot.Date } else { 'N/A' }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Software Inventory - $Computer - $anchor</title>
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
  .badge-removed { display: inline-block; background: #c62828; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
  .badge-update { display: inline-block; background: #1565c0; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
  .search-box { margin-bottom: 10px; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; width: 300px; max-width: 100%; box-sizing: border-box; }
  .search-box:focus { outline: none; border-color: #1a3a5c; box-shadow: 0 0 4px rgba(26,58,92,.3); }
  a.back-link { color: #1a3a5c; text-decoration: none; }
  a.back-link:hover { text-decoration: underline; }
  .theme-toggle { float: right; background: none; border: 1px solid #1a3a5c; color: #1a3a5c; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 13px; }
  .theme-toggle:hover { background: #1a3a5c; color: #fff; }
  @media (prefers-color-scheme: dark) {
    body.theme-auto { background: #1a1a2e; color: #e0e0e0; }
    body.theme-auto h1, body.theme-auto h2, body.theme-auto h3 { color: #80b0e0; }
    body.theme-auto .summary { background: #16213e; }
    body.theme-auto .summary-item { background: #0f3460; }
    body.theme-auto .summary-item .number { color: #80b0e0; }
    body.theme-auto .summary-item .label { color: #a0c0e0; }
    body.theme-auto .meta { color: #888; }
    body.theme-auto table { background: #16213e; }
    body.theme-auto th { background: #0f3460; }
    body.theme-auto th:hover { background: #1a4a7a; }
    body.theme-auto td { border-bottom: 1px solid #2a3a5e; }
    body.theme-auto tr:hover td { background: #1a2a4e; }
    body.theme-auto .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
    body.theme-auto .search-box:focus { border-color: #80b0e0; }
  }
  body.dark { background: #1a1a2e; color: #e0e0e0; }
  body.dark h1, body.dark h2, body.dark h3 { color: #80b0e0; }
  body.dark .summary { background: #16213e; }
  body.dark .summary-item { background: #0f3460; }
  body.dark .summary-item .number { color: #80b0e0; }
  body.dark .summary-item .label { color: #a0c0e0; }
  body.dark .meta { color: #888; }
  body.dark table { background: #16213e; }
  body.dark th { background: #0f3460; }
  body.dark th:hover { background: #1a4a7a; }
  body.dark td { border-bottom: 1px solid #2a3a5e; }
  body.dark tr:hover td { background: #1a2a4e; }
  body.dark .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
  body.dark .search-box:focus { border-color: #80b0e0; }
</style>
<script>
function toggleTheme() {
  var body = document.body;
  if (body.classList.contains('dark')) {
    body.classList.remove('dark');
    localStorage.setItem('theme', 'light');
  } else {
    body.classList.add('dark');
    localStorage.setItem('theme', 'dark');
  }
}
(function() {
  var saved = localStorage.getItem('theme');
  if (saved === 'dark') document.body.classList.add('dark');
  if (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches) document.body.classList.add('theme-auto');
})();
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
<a class="back-link" href="../../../index.html">&larr; Back to archive</a>
<h1>Software Inventory Report</h1>
<div class="summary">
  <div class="summary-grid">
    <div class="summary-item"><div class="number">$totalSw</div><div class="label">Installed Software</div></div>
    <div class="summary-item"><div class="number">$totalUp</div><div class="label">Installed Updates</div></div>
    <div class="summary-item"><div class="number"><span class="badge-new">+$newSwCount</span></div><div class="label">New Software</div></div>
    <div class="summary-item"><div class="number"><span class="badge-removed">$remSwCount</span></div><div class="label">Removed Software</div></div>
    <div class="summary-item"><div class="number"><span class="badge-update">+$newUpCount</span></div><div class="label">New Updates</div></div>
  </div>
  <div class="meta">Computer: <strong>$Computer</strong> &nbsp;|&nbsp; Generated: $reportDate &nbsp;|&nbsp; Previous snapshot: $prevDate</div>
</div>
"@

    # New Software section
    if ($newSwCount -gt 0) {
        $html += @"
<h2 class="section-title">New Software <span class="badge-new">$newSwCount</span></h2>
<table id="newSw-table"><thead><tr><th onclick="sortTable('newSw-table',0)">Name</th><th onclick="sortTable('newSw-table',1)">Version</th><th onclick="sortTable('newSw-table',2)">Install Date</th></tr></thead><tbody>$newSwRows</tbody></table>
"@
    }

    # Removed Software section
    if ($remSwCount -gt 0) {
        $html += @"
<h2 class="section-title">Removed Software <span class="badge-removed">$remSwCount</span></h2>
<table id="remSw-table"><thead><tr><th onclick="sortTable('remSw-table',0)">Name</th><th onclick="sortTable('remSw-table',1)">Version</th></tr></thead><tbody>$remSwRows</tbody></table>
"@
    }

    # New Updates section
    if ($newUpCount -gt 0) {
        $html += @"
<h2 class="section-title">New Updates <span class="badge-update">$newUpCount</span></h2>
<table id="newUp-table"><thead><tr><th onclick="sortTable('newUp-table',0)">Title</th><th onclick="sortTable('newUp-table',1)">Install Date</th></tr></thead><tbody>$newUpRows</tbody></table>
"@
    }

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
    $compFolder = $Computer -replace '[/\\:*?"<>|]', '_'
    $outDir = [System.IO.Path]::Combine($OutputDir, $compFolder, $year, $month)
    if (-not (Test-Path $outDir)) {
        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
    }

    $reportFile = Join-Path $outDir "report.html"
    $html | Out-File -FilePath $reportFile -Encoding utf8
    Write-Host "  Report saved: $reportFile"
    $reportFile
}

# ---------------------------------------------------------------
# New-CsvExport
# Exports software and updates data to CSV files alongside report output.
# ---------------------------------------------------------------
function New-CsvExport {
    param(
        [string]$Computer,
        [PSObject[]]$Software,
        [PSObject[]]$Updates,
        [string]$OutputDir
    )

    $now = Get-Date
    $year = $now.ToString('yyyy')
    $month = $now.ToString('MM')
    $timestamp = $now.ToString('yyyyMMdd-HHmm')

    $compFolder = $Computer -replace '[/\\:*?"<>|]', '_'
    $outDir = [System.IO.Path]::Combine($OutputDir, $compFolder, $year, $month)
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

# ---------------------------------------------------------------
# New-MonthReportHtml
# Generates a combined per-month index page at Output\YYYY\MM\index.html
# that aggregates all computers' snapshots with a Hostname column.
# ---------------------------------------------------------------
function New-MonthReportHtml {
    param(
        [string]$Year,
        [string]$Month,
        [string]$OutputDir,
        [string]$HistoryRoot
    )

    $monthDir = [System.IO.Path]::Combine($OutputDir, $Year, $Month)
    if (-not (Test-Path $monthDir)) {
        New-Item -Path $monthDir -ItemType Directory -Force | Out-Null
    }

    $monthName = switch ($Month) {
        '01' { 'January' }; '02' { 'February' }; '03' { 'March' }; '04' { 'April' }
        '05' { 'May' }; '06' { 'June' }; '07' { 'July' }; '08' { 'August' }
        '09' { 'September' }; '10' { 'October' }; '11' { 'November' }; '12' { 'December' }
        default { $Month }
    }

    # Scan all computer history folders for snapshots in this year/month
    $computerDirs = Get-ChildItem -Path $HistoryRoot -Directory -ErrorAction SilentlyContinue

    $allSoftware = @()
    $allUpdates = @()
    $computersFound = @()

    foreach ($compDir in $computerDirs) {
        $compFolder = $compDir.Name
        $snapDir = [System.IO.Path]::Combine($HistoryRoot, $compFolder, $Year, $Month)
        if (-not (Test-Path $snapDir)) { continue }

        $snapFiles = Get-ChildItem -Path $snapDir -Filter 'snapshot-*.json' |
            Sort-Object LastWriteTime -Descending
        if ($snapFiles.Count -eq 0) { continue }

        try {
            $snap = Get-Content -Path $snapFiles[0].FullName -Raw | ConvertFrom-Json
            $computersFound += $snap.Computer

            foreach ($sw in $snap.Software) {
                $allSoftware += [PSCustomObject]@{
                    Hostname    = $snap.Computer
                    Name        = $sw.Name
                    Version     = $sw.Version
                    Publisher   = $sw.Publisher
                    InstallDate = if ($sw.InstallDate) { $sw.InstallDate } else { 'Unknown' }
                }
            }
            foreach ($up in $snap.Updates) {
                $allUpdates += [PSCustomObject]@{
                    Hostname    = $snap.Computer
                    Title       = $up.Title
                    InstallDate = if ($up.InstallDate) { $up.InstallDate } else { 'Unknown' }
                }
            }
        } catch {
            Write-Warning "  Could not load snapshot for $compFolder : $_"
        }
    }

    $computersFound = $computersFound | Select-Object -Unique
    $compCount = $computersFound.Count
    $totalSw = $allSoftware.Count

    # Filter updates to only those installed in the selected month
    $allUpdates = $allUpdates | Where-Object {
        $_.InstallDate -ne 'Unknown' -and $_.InstallDate -match "^\d{4}-\d{2}-\d{2}$" -and
        $_.InstallDate -ge "$Year-$Month-01" -and
        $_.InstallDate -le "$Year-$Month-31"
    }
    $totalUp = $allUpdates.Count

    # Build software rows with hostname column
    $swRows = ''
    foreach ($item in $allSoftware | Sort-Object Hostname, Name) {
        $date = if ($item.InstallDate -and $item.InstallDate -ne 'Unknown') { $item.InstallDate } else { 'Unknown' }
        $swRows += @"
<tr><td>$(ConvertTo-HtmlEncoded $item.Hostname)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Name)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Version)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Publisher)</td>
    <td>$(ConvertTo-HtmlEncoded $date)</td></tr>
"@
    }

    # Build update rows with hostname column
    $upRows = ''
    foreach ($item in $allUpdates | Sort-Object Hostname, InstallDate -Descending) {
        $date = if ($item.InstallDate -and $item.InstallDate -ne 'Unknown') { $item.InstallDate } else { 'Unknown' }
        $title = $item.Title
        $upRows += @"
<tr><td>$(ConvertTo-HtmlEncoded $item.Hostname)</td>
    <td>$(ConvertTo-HtmlEncoded $title)</td>
    <td>$(ConvertTo-HtmlEncoded $date)</td></tr>
"@
    }

    # Computer list for summary
    $compList = ($computersFound | Sort-Object) -join ', '

    $now = Get-Date
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Software Inventory - $monthName $Year</title>
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
  .search-box { margin-bottom: 10px; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; width: 300px; max-width: 100%; box-sizing: border-box; }
  .search-box:focus { outline: none; border-color: #1a3a5c; box-shadow: 0 0 4px rgba(26,58,92,.3); }
  a.back-link { color: #1a3a5c; text-decoration: none; }
  a.back-link:hover { text-decoration: underline; }
  .theme-toggle { float: right; background: none; border: 1px solid #1a3a5c; color: #1a3a5c; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 13px; }
  .theme-toggle:hover { background: #1a3a5c; color: #fff; }
  @media (prefers-color-scheme: dark) {
    body.theme-auto { background: #1a1a2e; color: #e0e0e0; }
    body.theme-auto h1, body.theme-auto h2, body.theme-auto h3 { color: #80b0e0; }
    body.theme-auto .summary { background: #16213e; }
    body.theme-auto .summary-item { background: #0f3460; }
    body.theme-auto .summary-item .number { color: #80b0e0; }
    body.theme-auto .summary-item .label { color: #a0c0e0; }
    body.theme-auto .meta { color: #888; }
    body.theme-auto table { background: #16213e; }
    body.theme-auto th { background: #0f3460; }
    body.theme-auto th:hover { background: #1a4a7a; }
    body.theme-auto td { border-bottom: 1px solid #2a3a5e; }
    body.theme-auto tr:hover td { background: #1a2a4e; }
    body.theme-auto .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
    body.theme-auto .search-box:focus { border-color: #80b0e0; }
  }
  body.dark { background: #1a1a2e; color: #e0e0e0; }
  body.dark h1, body.dark h2, body.dark h3 { color: #80b0e0; }
  body.dark .summary { background: #16213e; }
  body.dark .summary-item { background: #0f3460; }
  body.dark .summary-item .number { color: #80b0e0; }
  body.dark .summary-item .label { color: #a0c0e0; }
  body.dark .meta { color: #888; }
  body.dark table { background: #16213e; }
  body.dark th { background: #0f3460; }
  body.dark th:hover { background: #1a4a7a; }
  body.dark td { border-bottom: 1px solid #2a3a5e; }
  body.dark tr:hover td { background: #1a2a4e; }
  body.dark .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
  body.dark .search-box:focus { border-color: #80b0e0; }
</style>
<script>
function toggleTheme() {
  var body = document.body;
  if (body.classList.contains('dark')) {
    body.classList.remove('dark');
    localStorage.setItem('theme', 'light');
  } else {
    body.classList.add('dark');
    localStorage.setItem('theme', 'dark');
  }
}
(function() {
  var saved = localStorage.getItem('theme');
  if (saved === 'dark') document.body.classList.add('dark');
  if (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches) document.body.classList.add('theme-auto');
})();
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
<a class="back-link" href="../../index.html">&larr; Back to archive</a>
<h1>$monthName $Year &mdash; Combined Inventory</h1>
<div class="summary">
  <div class="summary-grid">
    <div class="summary-item"><div class="number">$compCount</div><div class="label">Computers</div></div>
    <div class="summary-item"><div class="number">$totalSw</div><div class="label">Software Entries</div></div>
    <div class="summary-item"><div class="number">$totalUp</div><div class="label">Updates</div></div>
  </div>
  <div class="meta">Computers: $compList &nbsp;|&nbsp; Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss'))</div>
</div>

<h2 class="section-title">3rd Party Software</h2>
<input type="text" id="sw-filter" class="search-box" placeholder="Filter software..." onkeyup="filterTable('sw-filter','sw-table')">
<table id="sw-table"><thead><tr><th onclick="sortTable('sw-table',0)">Hostname</th><th onclick="sortTable('sw-table',1)">Name</th><th onclick="sortTable('sw-table',2)">Version</th><th onclick="sortTable('sw-table',3)">Publisher</th><th onclick="sortTable('sw-table',4)">Install Date</th></tr></thead><tbody>$swRows</tbody></table>

<h2 class="section-title">Windows Patches</h2>
<input type="text" id="up-filter" class="search-box" placeholder="Filter updates..." onkeyup="filterTable('up-filter','up-table')">
<table id="up-table"><thead><tr><th onclick="sortTable('up-table',0)">Hostname</th><th onclick="sortTable('up-table',1)">Title</th><th onclick="sortTable('up-table',2)">Install Date</th></tr></thead><tbody>$upRows</tbody></table>
</body></html>
"@

    $indexFile = Join-Path $monthDir "index.html"
    $html | Out-File -FilePath $indexFile -Encoding utf8
    Write-Host "  Month index saved: $indexFile"
}

# ---------------------------------------------------------------
# New-AllSoftwareHtml
# Generates Output\all-software.html aggregating every software
# entry from all historical snapshots, deduplicated by name.
# ---------------------------------------------------------------
function New-AllSoftwareHtml {
    param(
        [string]$HistoryRoot,
        [string]$OutputDir
    )

    $snapshotFiles = Get-ChildItem -Path $HistoryRoot -Recurse -Filter 'snapshot-*.json' -ErrorAction SilentlyContinue
    if ($snapshotFiles.Count -eq 0) { return }

    $allSoftware = @()
    foreach ($file in $snapshotFiles) {
        try {
            $snap = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
            $snapDate = $file.BaseName -replace '^snapshot-', ''
            foreach ($sw in $snap.Software) {
                $allSoftware += [PSCustomObject]@{
                    Name      = $sw.Name
                    Version   = $sw.Version
                    Publisher = $sw.Publisher
                    Computer  = $snap.Computer
                    SnapDate  = $snapDate
                }
            }
        } catch {
            # skip corrupt snapshots
        }
    }

    if ($allSoftware.Count -eq 0) { return }

    $grouped = $allSoftware | Group-Object -Property { ($_.Name -replace '\s+', ' ').Trim().ToLower() }

    $entries = @()
    foreach ($g in $grouped) {
        $items = $g.Group
        $computers = ($items | ForEach-Object { $_.Computer } | Select-Object -Unique | Sort-Object) -join ', '
        $compCount = ($items | ForEach-Object { $_.Computer } | Select-Object -Unique).Count
        $latest = $items | Sort-Object { $_.Version -eq 'Unknown' }, { $_.SnapDate } -Descending |
            Select-Object -First 1
        $entries += [PSCustomObject]@{
            Name          = $items[0].Name
            Version       = $latest.Version
            Publisher     = $latest.Publisher
            ComputerList  = $computers
            ComputerCount = $compCount
        }
    }

    $entries = $entries | Sort-Object Name
    $totalSw = $entries.Count

    $swRows = ''
    foreach ($item in $entries) {
        $swRows += @"
<tr><td>$(ConvertTo-HtmlEncoded $item.Name)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Version)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Publisher)</td>
    <td>$(ConvertTo-HtmlEncoded $item.ComputerCount)</td>
    <td>$(ConvertTo-HtmlEncoded $item.ComputerList)</td></tr>
"@
    }

    $now = Get-Date
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Software - Historical Inventory</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
  h1 { color: #1a3a5c; border-bottom: 2px solid #1a3a5c; padding-bottom: 8px; }
  .summary { background: #fff; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 20px; }
  .meta { font-size: 13px; color: #888; margin-top: 10px; }
  table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 25px; }
  th { background: #1a3a5c; color: #fff; padding: 10px 12px; text-align: left; font-weight: 600; cursor: pointer; }
  th:hover { background: #2a5a8c; }
  td { padding: 8px 12px; border-bottom: 1px solid #e0e0e0; word-break: break-word; }
  tr:hover td { background: #f0f5ff; }
  .search-box { margin-bottom: 10px; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; width: 300px; max-width: 100%; box-sizing: border-box; }
  .search-box:focus { outline: none; border-color: #1a3a5c; box-shadow: 0 0 4px rgba(26,58,92,.3); }
  a.back-link { color: #1a3a5c; text-decoration: none; }
  a.back-link:hover { text-decoration: underline; }
  .theme-toggle { float: right; background: none; border: 1px solid #1a3a5c; color: #1a3a5c; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 13px; }
  .theme-toggle:hover { background: #1a3a5c; color: #fff; }
  @media (prefers-color-scheme: dark) {
    body.theme-auto { background: #1a1a2e; color: #e0e0e0; }
    body.theme-auto h1 { color: #80b0e0; }
    body.theme-auto .summary { background: #16213e; }
    body.theme-auto table { background: #16213e; }
    body.theme-auto th { background: #0f3460; }
    body.theme-auto td { border-bottom: 1px solid #2a2a4e; }
    body.theme-auto tr:hover td { background: #1a2a4e; }
    body.theme-auto .meta { color: #888; }
  }
  body.dark { background: #1a1a2e; color: #e0e0e0; }
  body.dark h1 { color: #80b0e0; }
  body.dark .summary { background: #16213e; }
  body.dark table { background: #16213e; }
  body.dark th { background: #0f3460; }
  body.dark td { border-bottom: 1px solid #2a2a4e; }
  body.dark tr:hover td { background: #1a2a4e; }
  body.dark .meta { color: #888; }
</style>
<script>
function toggleTheme() {
  var body = document.body;
  if (body.classList.contains('dark')) { body.classList.remove('dark'); localStorage.setItem('theme', 'light'); }
  else { body.classList.add('dark'); localStorage.setItem('theme', 'dark'); }
}
(function() {
  var saved = localStorage.getItem('theme');
  if (saved === 'dark') document.body.classList.add('dark');
  if (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches) document.body.classList.add('theme-auto');
})();
function filterTable(inputId, tableId) {
  var input = document.getElementById(inputId);
  var filter = input.value.toUpperCase();
  var table = document.getElementById(tableId);
  var rows = table.getElementsByTagName('tr');
  for (var i = 1; i < rows.length; i++) {
    var cells = rows[i].getElementsByTagName('td');
    var show = false;
    for (var j = 0; j < cells.length; j++) {
      if (cells[j].textContent.toUpperCase().indexOf(filter) > -1) { show = true; break; }
    }
    rows[i].style.display = show ? '' : 'none';
  }
}
function sortTable(tableId, col) {
  var table = document.getElementById(tableId);
  var switching = true;
  var dir = 'asc';
  while (switching) {
    switching = false;
    var rows = table.rows;
    for (var i = 1; i < rows.length - 1; i++) {
      var x = rows[i].getElementsByTagName('td')[col];
      var y = rows[i + 1].getElementsByTagName('td')[col];
      var cmp = dir === 'asc' ? x.textContent.localeCompare(y.textContent) : y.textContent.localeCompare(x.textContent);
      if (cmp > 0) { rows[i].parentNode.insertBefore(rows[i + 1], rows[i]); switching = true; break; }
    }
    if (!switching && dir === 'asc') { dir = 'desc'; switching = true; }
  }
}
</script>
</head>
<body>
<button class="theme-toggle" onclick="toggleTheme()">&#9681; Theme</button>
<h1>All Software (Historical)</h1>
<div class="summary">
  <strong>$totalSw</strong> unique software titles found across all historical snapshots.
  <div class="meta">Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss')) &nbsp;|&nbsp; <a href="index.html" class="back-link">&larr; Back to Archive</a></div>
</div>
<input type="text" id="sw-filter" class="search-box" placeholder="Filter software..." onkeyup="filterTable('sw-filter','sw-table')">
<table id="sw-table"><thead><tr>
  <th onclick="sortTable('sw-table',0)">Name</th>
  <th onclick="sortTable('sw-table',1)">Version</th>
  <th onclick="sortTable('sw-table',2)">Publisher</th>
  <th onclick="sortTable('sw-table',3)">Computers</th>
  <th onclick="sortTable('sw-table',4)">Computer List</th>
</tr></thead><tbody>$swRows</tbody></table>
</body></html>
"@

    $outputFile = Join-Path $OutputDir "all-software.html"
    $html | Out-File -FilePath $outputFile -Encoding utf8
    Write-Host "  All Software page saved: $outputFile"
}

# ---------------------------------------------------------------
# New-WebsiteIndexHtml
# Generates the root Output\index.html with year/month navigation
# linking to combined month index pages across all computers.
# ---------------------------------------------------------------
function New-WebsiteIndexHtml {
    param(
        [string]$OutputDir,
        [int]$FailureCount = 0
    )

    if (-not (Test-Path $OutputDir)) { return }

    # Gather all report.html files across all computer subfolders
    $reports = Get-ChildItem -Path $OutputDir -Recurse -Filter 'report.html' -ErrorAction SilentlyContinue |
        Sort-Object FullName

    # Use relative paths from OutputDir so indices are predictable
    $OutputDirNorm = $OutputDir.TrimEnd('\/')

    # Group by year, then month
    $years = @{}
    foreach ($r in $reports) {
        $rel = $r.Directory.FullName.Substring($OutputDirNorm.Length + 1)
        $parts = $rel -split '[/\\]'
        if ($parts.Count -ge 3) {
            $compFolder = $parts[0]
            $y = $parts[1]
            $m = $parts[2]
            if (-not $years[$y]) { $years[$y] = @{} }
            if (-not $years[$y][$m]) { $years[$y][$m] = @() }
            $years[$y][$m] += @{ Computer = $compFolder; Path = $r.FullName }
        }
    }

    # If no report.html files found, try finding month index files
    if ($years.Count -eq 0) {
        $monthIndexes = Get-ChildItem -Path $OutputDir -Recurse -Filter 'index.html' -ErrorAction SilentlyContinue |
            Where-Object { $_.Directory.Parent.Name -match '^\d{4}$' } |
            Sort-Object FullName
        foreach ($r in $monthIndexes) {
            $rel = $r.Directory.FullName.Substring($OutputDirNorm.Length + 1)
            $parts = $rel -split '[/\\]'
            if ($parts.Count -ge 2) {
                $y = $parts[0]
                $m = $parts[1]
                if (-not $years[$y]) { $years[$y] = @{} }
                if (-not $years[$y][$m]) { $years[$y][$m] = @() }
                $years[$y][$m] += @{ Computer = ''; Path = $r.FullName }
            }
        }
    }

    # Build nav HTML
    $navHtml = ''
    $sortedYears = $years.Keys | Sort-Object -Descending
    foreach ($y in $sortedYears) {
        $navHtml += "<div class='year-group'><h2 class='year-heading'><a href='#year-$y'>$y</a></h2><div class='month-list'>`n"
        $sortedMonths = $years[$y].Keys | Sort-Object
        foreach ($m in $sortedMonths) {
            $monthName = switch ($m) {
                '01' { 'January' }; '02' { 'February' }; '03' { 'March' }; '04' { 'April' }
                '05' { 'May' }; '06' { 'June' }; '07' { 'July' }; '08' { 'August' }
                '09' { 'September' }; '10' { 'October' }; '11' { 'November' }; '12' { 'December' }
                default { $m }
            }
            $compCount = ($years[$y][$m] | ForEach-Object { $_.Computer } | Select-Object -Unique).Count
            $navHtml += "<a href='./$y/$m/index.html' class='month-link'>$monthName $y ($compCount computers)</a>`n"
        }
        $navHtml += "</div></div>"
    }

    $now = Get-Date
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Software Inventory Archive</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
  h1 { color: #1a3a5c; border-bottom: 2px solid #1a3a5c; padding-bottom: 8px; }
  .year-group { background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 15px; padding: 15px; }
  .year-heading { margin: 0 0 10px 0; }
  .year-heading a { color: #1a3a5c; text-decoration: none; }
  .year-heading a:hover { text-decoration: underline; }
  .month-list { display: flex; flex-wrap: wrap; gap: 8px; }
  .month-link { display: inline-block; background: #e8f0fe; color: #1a3a5c; padding: 8px 16px; border-radius: 4px; text-decoration: none; font-size: 14px; }
  .month-link:hover { background: #d0e0f0; }
  .year-section { margin-top: 30px; color: #1a3a5c; border-bottom: 1px solid #ccc; padding-bottom: 5px; }
  .month-section { margin-top: 20px; }
  .month-section a { color: #2a5a8c; text-decoration: none; }
  .month-section a:hover { text-decoration: underline; }
  .meta { font-size: 13px; color: #888; margin-top: 20px; }
  .failures-link { display: inline-block; padding: 8px 16px; border-radius: 4px; text-decoration: none; font-size: 14px; margin-top: 10px; }
  .failures-link.green { background: #e8f5e9; color: #2e7d32; }
  .failures-link.green:hover { background: #c8e6c9; }
  .failures-link.red { background: #ffebee; color: #c62828; }
  .failures-link.red:hover { background: #ffcdd2; }
  .theme-toggle { float: right; background: none; border: 1px solid #1a3a5c; color: #1a3a5c; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 13px; }
  .theme-toggle:hover { background: #1a3a5c; color: #fff; }
  @media (prefers-color-scheme: dark) {
    body.theme-auto { background: #1a1a2e; color: #e0e0e0; }
    body.theme-auto h1 { color: #80b0e0; }
    body.theme-auto .year-group { background: #16213e; }
    body.theme-auto .month-link { background: #0f3460; color: #80b0e0; }
    body.theme-auto .month-link:hover { background: #1a4a7a; }
    body.theme-auto .meta { color: #888; }
    body.theme-auto .failures-link.green { background: #1b5e20; color: #a5d6a7; }
    body.theme-auto .failures-link.red { background: #b71c1c; color: #ffcdd2; }
  }
  body.dark { background: #1a1a2e; color: #e0e0e0; }
  body.dark h1 { color: #80b0e0; }
  body.dark .year-group { background: #16213e; }
  body.dark .month-link { background: #0f3460; color: #80b0e0; }
  body.dark .month-link:hover { background: #1a4a7a; }
  body.dark .meta { color: #888; }
  body.dark .failures-link.green { background: #1b5e20; color: #a5d6a7; }
  body.dark .failures-link.red { background: #b71c1c; color: #ffcdd2; }
</style>
<script>
function toggleTheme() {
  var body = document.body;
  if (body.classList.contains('dark')) {
    body.classList.remove('dark');
    localStorage.setItem('theme', 'light');
  } else {
    body.classList.add('dark');
    localStorage.setItem('theme', 'dark');
  }
}
(function() {
  var saved = localStorage.getItem('theme');
  if (saved === 'dark') document.body.classList.add('dark');
  if (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches) document.body.classList.add('theme-auto');
})();
</script>
</head>
<body>
<button class="theme-toggle" onclick="toggleTheme()">&#9681; Theme</button>
<h1>Software Inventory Archive</h1>

<div class="year-group">
  <h2 style="margin:0 0 10px 0;">Jump to Year</h2>
  <div class="month-list">
$(
    ($sortedYears | ForEach-Object { "<a href='#year-$_' class='month-link'>$_</a>" }) -join "`n"
)
  </div>
  <div style="margin-top: 15px; border-top: 1px solid #e0e0e0; padding-top: 12px;">
    <a href="all-software.html" class="failures-link" style="background:#e8f0fe;color:#1a3a5c;">View All Software</a>
    <a href="failures.html" class="failures-link $(if ($FailureCount -gt 0) { 'red' } else { 'green' })">$(if ($FailureCount -gt 0) { "&#9888; View Failures ($FailureCount)" } else { "&#10003; No Failures" })</a>
  </div>
</div>

$navHtml

<div class="meta">Last updated: $($now.ToString('yyyy-MM-dd HH:mm:ss'))</div>
</body></html>
"@

    $indexFile = Join-Path $OutputDir "index.html"
    $html | Out-File -FilePath $indexFile -Encoding utf8
    Write-Host "  Root index saved: $indexFile"
}

# ---------------------------------------------------------------
# New-FailuresHtml
# Generates Output\failures.html listing computers that could not
# be inventoried, with error details and timestamps.
# ---------------------------------------------------------------
function New-FailuresHtml {
    param(
        [PSObject[]]$Failures,
        [string]$OutputDir
    )

    if (-not (Test-Path $OutputDir)) {
        New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
    }

    if (-not $Failures -or $Failures.Count -eq 0) {
        # No failures — generate a clean page stating that
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Inventory Failures</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
  h1 { color: #1a3a5c; border-bottom: 2px solid #1a3a5c; padding-bottom: 8px; }
  .success { background: #e8f5e9; padding: 30px; border-radius: 6px; text-align: center; font-size: 18px; color: #2e7d32; }
  a.back-link { color: #1a3a5c; text-decoration: none; }
  a.back-link:hover { text-decoration: underline; }
  .meta { font-size: 13px; color: #888; margin-top: 20px; }
  .theme-toggle { float: right; background: none; border: 1px solid #2e7d32; color: #2e7d32; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 13px; }
  .theme-toggle:hover { background: #2e7d32; color: #fff; }
  @media (prefers-color-scheme: dark) {
    body.theme-auto { background: #1a1a2e; color: #e0e0e0; }
    body.theme-auto h1 { color: #81c784; }
    body.theme-auto .success { background: #1b5e20; color: #a5d6a7; }
  }
  body.dark { background: #1a1a2e; color: #e0e0e0; }
  body.dark h1 { color: #81c784; }
  body.dark .success { background: #1b5e20; color: #a5d6a7; }
</style>
<script>
function toggleTheme() {
  var body = document.body;
  if (body.classList.contains('dark')) {
    body.classList.remove('dark');
    localStorage.setItem('theme', 'light');
  } else {
    body.classList.add('dark');
    localStorage.setItem('theme', 'dark');
  }
}
(function() {
  var saved = localStorage.getItem('theme');
  if (saved === 'dark') document.body.classList.add('dark');
  if (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches) document.body.classList.add('theme-auto');
})();
</script>
</head>
<body>
<button class="theme-toggle" onclick="toggleTheme()">&#9681; Theme</button>
<a class="back-link" href="index.html">&larr; Back to archive</a>
<h1>Inventory Failures</h1>
<div class="success">All computers were successfully inventoried.</div>
<div class="meta">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
</body></html>
"@
    } else {
        $rows = ''
        foreach ($f in $Failures) {
            $errors = $f.Errors -replace '`n', '<br>'
            $rows += @"
<tr><td>$(ConvertTo-HtmlEncoded $f.Computer)</td>
    <td>$(ConvertTo-HtmlEncoded $f.Timestamp)</td>
    <td>$errors</td></tr>
"@
        }

        $count = $Failures.Count
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Inventory Failures ($count)</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
  h1 { color: #c62828; border-bottom: 2px solid #c62828; padding-bottom: 8px; }
  .summary { background: #ffebee; padding: 15px; border-radius: 6px; margin-bottom: 20px; font-size: 15px; }
  table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,.1); }
  th { background: #c62828; color: #fff; padding: 10px 12px; text-align: left; font-weight: 600; cursor: pointer; }
  th:hover { background: #e53935; }
  td { padding: 10px 12px; border-bottom: 1px solid #e0e0e0; vertical-align: top; }
  tr:hover td { background: #fff5f5; }
  a.back-link { color: #1a3a5c; text-decoration: none; }
  a.back-link:hover { text-decoration: underline; }
  .meta { font-size: 13px; color: #888; margin-top: 20px; }
   .error-detail { font-family: 'Consolas', 'Courier New', monospace; font-size: 12px; white-space: pre-wrap; word-break: break-all; }
   .search-box { margin-bottom: 10px; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; width: 300px; max-width: 100%; box-sizing: border-box; }
   .search-box:focus { outline: none; border-color: #c62828; box-shadow: 0 0 4px rgba(198,40,40,.3); }
   .theme-toggle { float: right; background: none; border: 1px solid #c62828; color: #c62828; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-size: 13px; }
   .theme-toggle:hover { background: #c62828; color: #fff; }
   @media (prefers-color-scheme: dark) {
     body.theme-auto { background: #1a1a2e; color: #e0e0e0; }
     body.theme-auto h1 { color: #e57373; }
     body.theme-auto .summary { background: #16213e; }
     body.theme-auto table { background: #16213e; }
     body.theme-auto th { background: #b71c1c; }
     body.theme-auto th:hover { background: #c62828; }
     body.theme-auto td { border-bottom: 1px solid #2a3a5e; }
     body.theme-auto tr:hover td { background: #1a2a4e; }
     body.theme-auto .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
   }
   body.dark { background: #1a1a2e; color: #e0e0e0; }
   body.dark h1 { color: #e57373; }
   body.dark .summary { background: #16213e; }
   body.dark table { background: #16213e; }
   body.dark th { background: #b71c1c; }
   body.dark th:hover { background: #c62828; }
   body.dark td { border-bottom: 1px solid #2a3a5e; }
   body.dark tr:hover td { background: #1a2a4e; }
   body.dark .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
</style>
<script>
function toggleTheme() {
  var body = document.body;
  if (body.classList.contains('dark')) {
    body.classList.remove('dark');
    localStorage.setItem('theme', 'light');
  } else {
    body.classList.add('dark');
    localStorage.setItem('theme', 'dark');
  }
}
(function() {
  var saved = localStorage.getItem('theme');
  if (saved === 'dark') document.body.classList.add('dark');
  if (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches) document.body.classList.add('theme-auto');
})();
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
<a class="back-link" href="index.html">&larr; Back to archive</a>
<h1>Inventory Failures</h1>
<div class="summary"><strong>$count computer(s)</strong> could not be inventoried. See details below.</div>
<input type="text" id="fail-filter" class="search-box" placeholder="Filter failures..." onkeyup="filterTable('fail-filter','fail-table')">
<table id="fail-table"><thead><tr><th onclick="sortTable('fail-table',0)">Computer</th><th onclick="sortTable('fail-table',1)">Timestamp</th><th onclick="sortTable('fail-table',2)">Error Details</th></tr></thead><tbody>$rows</tbody></table>
<div class="meta">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>
</body></html>
"@
    }

    $failFile = Join-Path $OutputDir "failures.html"
    $html | Out-File -FilePath $failFile -Encoding utf8
    Write-Host "  Failures page saved: $failFile"
}

# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------

Write-Host "Software Inventory V7 - Starting"
Write-Host "Output path: $OutputPath"
Write-Host "History path: $HistoryPath"
Write-Host "Target computers: $($ComputerName -join ', ')"
if ($ThrottleLimit -ne 5) { Write-Host "Throttle limit: $ThrottleLimit" }
Write-Host ""

$failures = @()

# ---------------------------------------------------------------
# Split hosts into local vs remote for parallel collection
# ---------------------------------------------------------------
$localHosts   = @()
$remoteHosts  = @()
foreach ($computer in $ComputerName) {
    $computer = $computer.Trim()
    if (-not $computer) { continue }
    if (Test-IsLocalComputer $computer) { $localHosts += $computer }
    else { $remoteHosts += $computer }
}

# Health check remote hosts — filter out unreachable ones
$healthyRemote = @()
foreach ($computer in $remoteHosts) {
    if (Test-ComputerConnectivity -Computer $computer) {
        $healthyRemote += $computer
    } else {
        Write-Warning "  $computer unreachable via WinRM. Skipping."
        $failures += [PSCustomObject]@{
            Computer  = $computer
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            Errors    = "Unreachable via WinRM - host is offline or WinRM is not configured"
        }
    }
}

# ---------------------------------------------------------------
# PHASE 1 — Collect inventory from all hosts
# ---------------------------------------------------------------
$allResults = @()

# -- Local hosts: sequential (inline) --
foreach ($computer in $localHosts) {
    Write-Host "========================================"
    Write-Host "Processing (local): $computer"
    Write-Host "========================================"
    $sw = @(); $up = @()
    Write-Host "  Collecting software (registry)..."
    try { $sw = Get-SoftwareInventory -Computer $computer; Write-Host "    Found $($sw.Count) software entries" }
    catch { Write-Warning "    Software inventory failed for $computer : $_" }
    Write-Host "  Collecting updates (WUA)..."
    try { $up = Get-InstalledUpdates -Computer $computer; Write-Host "    Found $($up.Count) updates" }
    catch { Write-Warning "    Update query failed for $computer : $_" }
    $allResults += [PSCustomObject]@{ Computer = $computer; Software = $sw; Updates = $up }
}

# -- Remote hosts: parallel via Invoke-Command --
if ($healthyRemote.Count -gt 0) {
    Write-Host "========================================"
    Write-Host "Remote collection: $($healthyRemote.Count) host(s) (throttle=$ThrottleLimit)"
    Write-Host "========================================"

    $remoteScriptBlock = [ScriptBlock]::Create(@"
function Get-LocalSoftware {
`$(${function:Get-LocalSoftware})
}
function Get-LocalUpdates {
`$(${function:Get-LocalUpdates})
}
function Get-LocalHotfixFallback {
`$(${function:Get-LocalHotfixFallback})
}
function Merge-UpdateDuplicates {
`$(${function:Merge-UpdateDuplicates})
}
`$sw = @(); `$up = @()
try { `$sw = Get-LocalSoftware } catch { }
try { `$up = Get-LocalUpdates } catch { }
if (`$up) { `$up = Merge-UpdateDuplicates `$up }
@{ Software = `$sw; Updates = `$up }
"@)

    $remoteResults = Invoke-Command -ComputerName $healthyRemote -ScriptBlock $remoteScriptBlock `
        -ThrottleLimit $ThrottleLimit -ErrorAction SilentlyContinue -ErrorVariable remoteErrors

    foreach ($r in $remoteResults) {
        $allResults += [PSCustomObject]@{
            Computer = $r.PSComputerName
            Software = $r.Software
            Updates  = $r.Updates
        }
    }
    if ($remoteErrors) {
        foreach ($e in $remoteErrors) {
            $failedComp = if ($e.TargetObject) { $e.TargetObject } else { 'Unknown' }
            Write-Warning "  Remote collection failed for $failedComp : $($e.Exception.Message)"
            $failures += [PSCustomObject]@{
                Computer  = $failedComp
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Errors    = "Remote collection failed: $($e.Exception.Message)"
            }
        }
    }
}

# ---------------------------------------------------------------
# PHASE 2 — Process all results sequentially
# ---------------------------------------------------------------
foreach ($result in $allResults) {
    $computer = $result.Computer
    $software = $result.Software
    $updates  = $result.Updates

    Write-Host "========================================"
    Write-Host "Processing: $computer"
    Write-Host "========================================"

    # Check if computer failed completely
    if ($software.Count -eq 0 -and $updates.Count -eq 0) {
        Write-Warning "  $computer failed — no data collected. Skipping snapshot and report."
        $failures += [PSCustomObject]@{
            Computer  = $computer
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            Errors    = "No data collected — verify connectivity and credentials"
        }
        Write-Host ""
        continue
    }

    # 2. Load previous snapshot
    Write-Host "  Loading previous snapshot..."
    $prevSnapshot = Load-HistorySnapshot -Computer $computer -HistoryRoot $HistoryPath

    # 3. Compare snapshots
    Write-Host "  Comparing snapshots..."
    $comparison = Compare-Snapshots -CurrentSoftware $software -CurrentUpdates $updates -PreviousSnapshot $prevSnapshot
    Write-Host "    New software: $($comparison.NewSoftware.Count)"
    Write-Host "    Removed software: $($comparison.RemovedSoftware.Count)"
    Write-Host "    New updates: $($comparison.NewUpdates.Count)"

    # 4. Save current snapshot
    Write-Host "  Saving snapshot..."
    $snapshotFile = Save-HistorySnapshot -Computer $computer -Software $software -Updates $updates -HistoryRoot $HistoryPath

    # 5. Generate HTML report
    Write-Host "  Generating HTML report..."
    $reportFile = New-InventoryHtmlReport -Computer $computer -Software $software -Updates $updates `
        -Comparison $comparison -PreviousSnapshot $prevSnapshot -OutputDir $OutputPath

    # 5b. Export CSV (optional)
    if ($ExportCsv) {
        Write-Host "  Exporting CSV..."
        New-CsvExport -Computer $computer -Software $software -Updates $updates -OutputDir $OutputPath
    }

    # 6. PassThru
    if ($PassThru) {
        [PSCustomObject]@{
            Computer         = $computer
            Software         = $software
            Updates          = $updates
            Comparison       = $comparison
            SnapshotFile     = $snapshotFile
            ReportFile       = $reportFile
        }
    }

    Write-Host ""
}

# Generate combined month page, failures page, and root index
$now = Get-Date
$currentYear = $now.ToString('yyyy')
$currentMonth = $now.ToString('MM')
Write-Host "Generating combined month report..."
New-MonthReportHtml -Year $currentYear -Month $currentMonth -OutputDir $OutputPath -HistoryRoot $HistoryPath
Write-Host "Generating failures page..."
New-FailuresHtml -Failures $failures -OutputDir $OutputPath
Write-Host "Generating All Software page..."
New-AllSoftwareHtml -HistoryRoot $HistoryPath -OutputDir $OutputPath
Write-Host "Generating root website index..."
New-WebsiteIndexHtml -OutputDir $OutputPath -FailureCount $failures.Count

Write-Host "Done."
