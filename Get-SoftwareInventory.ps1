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
<a class="back-link" href="../../../index.html">&larr; Back to archive</a>
<h1>Software Inventory Report</h1>
<div class="summary">
  <div class="summary-grid">
    <div class="summary-item"><div class="number"><a href="../../../all-software.html" style="color:inherit;text-decoration:none">$totalSw</a></div><div class="label">3rd Party Software</div></div>
    <div class="summary-item"><div class="number"><a href="../../../all-software.html" style="color:inherit;text-decoration:none">$totalUp</a></div><div class="label">Windows Patches</div></div>
    <div class="summary-item"><div class="number"><span class="badge-new">+$newSwCount</span></div><div class="label">New Software</div></div>
    <div class="summary-item"><div class="number"><span class="badge-removed">$remSwCount</span></div><div class="label">Removed Software</div></div>
    <div class="summary-item"><div class="number"><span class="badge-update">+$newUpCount</span></div><div class="label">New Patches</div></div>
  </div>
  <div class="meta">Computer: <strong>$Computer</strong> &nbsp;|&nbsp; Generated: $reportDate &nbsp;|&nbsp; Previous snapshot: $prevDate</div>
</div>
"@

    # New Software section
    $html += @"
<h2 class="section-title">New Software <span class="badge-new">$newSwCount</span></h2>
<table id="newSw-table"><thead><tr><th onclick="sortTable('newSw-table',0)">Name</th><th onclick="sortTable('newSw-table',1)">Version</th><th onclick="sortTable('newSw-table',2)">Install Date</th></tr></thead><tbody>$newSwRows</tbody></table>
"@

    # Removed Software section
    if ($remSwCount -gt 0) {
        $html += @"
<h2 class="section-title">Removed Software <span class="badge-removed">$remSwCount</span></h2>
<table id="remSw-table"><thead><tr><th onclick="sortTable('remSw-table',0)">Name</th><th onclick="sortTable('remSw-table',1)">Version</th></tr></thead><tbody>$remSwRows</tbody></table>
"@
    }

    # New Updates section
    $html += @"
<h2 class="section-title">New Patches <span class="badge-update">$newUpCount</span></h2>
<table id="newUp-table"><thead><tr><th onclick="sortTable('newUp-table',0)">Title</th><th onclick="sortTable('newUp-table',1)">Install Date</th></tr></thead><tbody>$newUpRows</tbody></table>
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

    # Filter software to only those installed in the selected month
    $allSoftware = $allSoftware | Where-Object {
        try { $dt = [datetime]$_.InstallDate; $dt.Year -eq [int]$Year -and $dt.Month -eq [int]$Month }
        catch { $false }
    }
    $totalSw = $allSoftware.Count

    # Filter updates to only those installed in the selected month
    $allUpdates = $allUpdates | Where-Object {
        try { $dt = [datetime]$_.InstallDate; $dt.Year -eq [int]$Year -and $dt.Month -eq [int]$Month }
        catch { $false }
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
<a class="back-link" href="../../index.html">&larr; Back to archive</a>
<h1>$monthName $Year &mdash; Combined Inventory</h1>
<div class="summary">
  <div class="summary-grid">
    <div class="summary-item"><div class="number"><a href="../../computers.html" style="color:inherit;text-decoration:none">$compCount</a></div><div class="label">Computers</div></div>
    <div class="summary-item"><div class="number"><a href="../../all-software.html" style="color:inherit;text-decoration:none">$totalSw</a></div><div class="label">3rd Party Software</div></div>
    <div class="summary-item"><div class="number"><a href="../../all-software.html" style="color:inherit;text-decoration:none">$totalUp</a></div><div class="label">Windows Patches</div></div>
  </div>
  <div class="meta">Computers: $compList &nbsp;|&nbsp; Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss'))</div>
</div>

<h2 class="section-title">3rd Party Software <span class="badge-new">$totalSw</span></h2>
<input type="text" id="sw-filter" class="search-box" placeholder="Filter software..." onkeyup="filterTable('sw-filter','sw-table')">
<table id="sw-table"><thead><tr><th onclick="sortTable('sw-table',0)">Hostname</th><th onclick="sortTable('sw-table',1)">Name</th><th onclick="sortTable('sw-table',2)">Version</th><th onclick="sortTable('sw-table',3)">Publisher</th><th onclick="sortTable('sw-table',4)">Install Date</th></tr></thead><tbody>$swRows</tbody></table>

<h2 class="section-title">Windows Patches <span class="badge-update">$totalUp</span></h2>
<input type="text" id="up-filter" class="search-box" placeholder="Filter updates..." onkeyup="filterTable('up-filter','up-table')">
<table id="up-table"><thead><tr><th onclick="sortTable('up-table',0)">Hostname</th><th onclick="sortTable('up-table',1)">Title</th><th onclick="sortTable('up-table',2)">Install Date</th></tr></thead><tbody>$upRows</tbody></table>
</body></html>
"@
    $indexFile = Join-Path $monthDir "index.html"
    $html | Out-File -FilePath $indexFile -Encoding utf8
    Write-Host "  Month index saved: $indexFile ($totalSw sw, $totalUp patches)"
}

# ---------------------------------------------------------------
# New-AllSoftwareHtml
# Generates Output\all-software.html aggregating every software
# entry and patch from all historical snapshots, deduplicated.
# ---------------------------------------------------------------
function New-AllSoftwareHtml {
    param(
        [string]$HistoryRoot,
        [string]$OutputDir
    )

    $snapshotFiles = Get-ChildItem -Path $HistoryRoot -Recurse -Filter 'snapshot-*.json' -ErrorAction SilentlyContinue
    if ($snapshotFiles.Count -eq 0) { return }

    [System.Collections.ArrayList]$allSoftware = @()
    [System.Collections.ArrayList]$allPatches = @()
    $computerLatest = @{}
    foreach ($file in $snapshotFiles) {
        try {
            $snap = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
            $snapDate = $file.BaseName -replace '^snapshot-', ''
            $compFolder = $snap.Computer -replace '[/\\:*?"<>|]', '_'
            $snapYear = $snapDate.Substring(0, 4)
            $snapMonth = $snapDate.Substring(4, 2)
            if (-not $computerLatest.ContainsKey($compFolder) -or $snapDate -gt $computerLatest[$compFolder].SnapDate) {
                $computerLatest[$compFolder] = @{ SnapDate = $snapDate; Year = $snapYear; Month = $snapMonth; Computer = $snap.Computer }
            }
            foreach ($sw in $snap.Software) {
                $null = $allSoftware.Add([PSCustomObject]@{
                    Name        = $sw.Name
                    Version     = $sw.Version
                    Publisher   = $sw.Publisher
                    InstallDate = if ($sw.InstallDate) { $sw.InstallDate } else { 'Unknown' }
                    Computer    = $snap.Computer
                    SnapDate    = $snapDate
                })
            }
            foreach ($up in $snap.Updates) {
                $null = $allPatches.Add([PSCustomObject]@{
                    Title       = $up.Title
                    InstallDate = if ($up.InstallDate) { $up.InstallDate } else { 'Unknown' }
                    Computer    = $snap.Computer
                })
            }
        } catch {
            # skip corrupt snapshots
        }
    }

    # --- Software dedup ---
    $swEntries = @()
    if ($allSoftware.Count -gt 0) {
        $grouped = $allSoftware | Group-Object -Property { ($_.Name -replace '\s+', ' ').Trim().ToLower() }
        foreach ($g in $grouped) {
            $items = $g.Group
            $compNames = $items | ForEach-Object { $_.Computer } | Select-Object -Unique | Sort-Object
            $compCount = $compNames.Count
            $compLinks = @{}
            foreach ($cn in $compNames) {
                $cf = $cn -replace '[/\\:*?"<>|]', '_'
                if ($computerLatest.ContainsKey($cf)) {
                    $cl = $computerLatest[$cf]
                    $compLinks[$cn] = "<a href='./$cf/$($cl.Year)/$($cl.Month)/report.html'>$(ConvertTo-HtmlEncoded $cn)</a>"
                } else {
                    $compLinks[$cn] = $(ConvertTo-HtmlEncoded $cn)
                }
            }
            $computers = ($compNames | ForEach-Object { $compLinks[$_] }) -join ', '
            $latest = $items | Sort-Object { $_.Version -eq 'Unknown' }, { $_.SnapDate } -Descending |
                Select-Object -First 1
            $latestInstallDate = $items | Sort-Object { if ($_.InstallDate -eq 'Unknown') { '0000-00-00' } else { $_.InstallDate } } -Descending | Select-Object -First 1
            $swEntries += [PSCustomObject]@{
                Name          = $items[0].Name
                Version       = $latest.Version
                Publisher     = $latest.Publisher
                InstallDate   = $latestInstallDate.InstallDate
                ComputerList  = $computers
                ComputerCount = $compCount
            }
        }
        $swEntries = $swEntries | Sort-Object Name
    }
    $totalSw = $swEntries.Count

    $swSb = New-Object System.Text.StringBuilder
    foreach ($item in $swEntries) {
        $null = $swSb.Append(@"
<tr><td>$(ConvertTo-HtmlEncoded $item.Name)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Version)</td>
    <td>$(ConvertTo-HtmlEncoded $item.Publisher)</td>
    <td>$(ConvertTo-HtmlEncoded $item.InstallDate)</td>
    <td><a href="computers.html">$($item.ComputerCount)</a></td>
    <td>$($item.ComputerList)</td></tr>
"@)
    }
    $swRows = $swSb.ToString()

    # --- Patch dedup ---
    $patchEntries = @()
    if ($allPatches.Count -gt 0) {
        $grouped = $allPatches | Group-Object -Property { ($_.Title -replace '\s+', ' ').Trim().ToLower() }
        foreach ($g in $grouped) {
            $items = $g.Group
            $compNames = $items | ForEach-Object { $_.Computer } | Select-Object -Unique | Sort-Object
            $compCount = $compNames.Count
            $compLinks = @{}
            foreach ($cn in $compNames) {
                $cf = $cn -replace '[/\\:*?"<>|]', '_'
                if ($computerLatest.ContainsKey($cf)) {
                    $cl = $computerLatest[$cf]
                    $compLinks[$cn] = "<a href='./$cf/$($cl.Year)/$($cl.Month)/report.html'>$(ConvertTo-HtmlEncoded $cn)</a>"
                } else {
                    $compLinks[$cn] = $(ConvertTo-HtmlEncoded $cn)
                }
            }
            $computers = ($compNames | ForEach-Object { $compLinks[$_] }) -join ', '
            $latest = $items | Sort-Object {
                if ($_.InstallDate -eq 'Unknown') { '0000-00-00' } else { $_.InstallDate }
            } -Descending | Select-Object -First 1
            $patchEntries += [PSCustomObject]@{
                Title         = $items[0].Title
                InstallDate   = $latest.InstallDate
                ComputerList  = $computers
                ComputerCount = $compCount
            }
        }
        $patchEntries = $patchEntries | Sort-Object Title
    }
    $totalPatches = $patchEntries.Count

    $patchSb = New-Object System.Text.StringBuilder
    foreach ($item in $patchEntries) {
        $null = $patchSb.Append(@"
<tr><td>$(ConvertTo-HtmlEncoded $item.Title)</td>
    <td><a href="computers.html">$($item.ComputerCount)</a></td>
    <td>$($item.ComputerList)</td>
    <td>$(ConvertTo-HtmlEncoded $item.InstallDate)</td></tr>
"@)
    }
    $patchRows = $patchSb.ToString()

    $now = Get-Date
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Software &amp; Patches - Global Inventory</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
  h1 { color: #1a3a5c; border-bottom: 2px solid #1a3a5c; padding-bottom: 8px; }
  .summary { background: #fff; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 20px; }
  .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px,1fr)); gap: 10px; margin-bottom: 15px; }
  .summary-item { background: #e8f0fe; padding: 10px; border-radius: 4px; text-align: center; }
  .summary-item .number { font-size: 24px; font-weight: bold; color: #1a3a5c; }
  .summary-item .label { font-size: 13px; color: #666; }
  .section-title { margin: 25px 0 10px 0; }
  .badge-new { display: inline-block; background: #2e7d32; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
  .badge-removed { display: inline-block; background: #c62828; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
  .badge-update { display: inline-block; background: #1565c0; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; }
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
    html.theme-auto body { background: #1a1a2e; color: #e0e0e0; }
    html.theme-auto body h1 { color: #80b0e0; }
    html.theme-auto body .summary { background: #16213e; }
    html.theme-auto body .summary-item { background: #0f3460; }
    html.theme-auto body .summary-item .number { color: #80b0e0; }
    html.theme-auto body table { background: #16213e; }
    html.theme-auto body th { background: #0f3460; }
    html.theme-auto body td { border-bottom: 1px solid #2a2a4e; }
    html.theme-auto body tr:hover td { background: #1a2a4e; }
    html.theme-auto body .meta { color: #888; }
    html.theme-auto body .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
    html.theme-auto body .search-box:focus { border-color: #80b0e0; }
  }
  html.dark body { background: #1a1a2e; color: #e0e0e0; }
  html.dark body h1 { color: #80b0e0; }
  html.dark body .summary { background: #16213e; }
  html.dark body .summary-item { background: #0f3460; }
  html.dark body .summary-item .number { color: #80b0e0; }
  html.dark body table { background: #16213e; }
  html.dark body th { background: #0f3460; }
  html.dark body td { border-bottom: 1px solid #2a2a4e; }
  html.dark body tr:hover td { background: #1a2a4e; }
  html.dark body .meta { color: #888; }
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
<h1>All Software &amp; Patches</h1>
<div class="summary">
  <div class="summary-grid">
    <div class="summary-item">
      <div class="number"><a href="#software-section" style="color:inherit;text-decoration:none">$totalSw</a></div>
      <div class="label">Unique Software Titles</div>
    </div>
    <div class="summary-item">
      <div class="number"><a href="#patches-section" style="color:inherit;text-decoration:none">$totalPatches</a></div>
      <div class="label">Unique Windows Patches</div>
    </div>
  </div>
  <div class="meta">Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss')) &nbsp;|&nbsp; <a href="index.html" class="back-link">&larr; Back to Archive</a></div>
</div>

<h2 class="section-title" id="software-section">3rd Party Software <span class="badge-new">$totalSw</span></h2>
<input type="text" id="sw-filter" class="search-box" placeholder="Filter software..." onkeyup="filterTable('sw-filter','sw-table')">
<table id="sw-table"><thead><tr>
  <th onclick="sortTable('sw-table',0)">Name</th>
  <th onclick="sortTable('sw-table',1)">Version</th>
  <th onclick="sortTable('sw-table',2)">Publisher</th>
  <th onclick="sortTable('sw-table',3)">Latest Install Date</th>
  <th onclick="sortTable('sw-table',4)">Computers</th>
  <th onclick="sortTable('sw-table',5)">Computer List</th>
</tr></thead><tbody>$swRows</tbody></table>

<h2 class="section-title" id="patches-section">Windows Patches <span class="badge-update">$totalPatches</span></h2>
<input type="text" id="patch-filter" class="search-box" placeholder="Filter patches..." onkeyup="filterTable('patch-filter','patch-table')">
<table id="patch-table"><thead><tr>
  <th onclick="sortTable('patch-table',0)">Title</th>
  <th onclick="sortTable('patch-table',1)">Computers</th>
  <th onclick="sortTable('patch-table',2)">Computer List</th>
  <th onclick="sortTable('patch-table',3)">Latest Install Date</th>
</tr></thead><tbody>$patchRows</tbody></table>
</body></html>
"@

    $outputFile = Join-Path $OutputDir "all-software.html"
    $html | Out-File -FilePath $outputFile -Encoding utf8
    Write-Host "  All Software & Patches page saved: $outputFile"
}

# ---------------------------------------------------------------
# New-ComputersHtml
# Generates Output\computers.html listing all unique hostnames
# from history, each linked to its per-computer report.
# ---------------------------------------------------------------
function New-ComputersHtml {
    param(
        [string]$HistoryRoot,
        [string]$OutputDir
    )

    $computerDirs = Get-ChildItem -Path $HistoryRoot -Directory -ErrorAction SilentlyContinue
    if ($computerDirs.Count -eq 0) { return }

    $entries = @()
    foreach ($compDir in $computerDirs) {
        $compFolder = $compDir.Name
        $yearDirs = Get-ChildItem -Path $compDir.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
        $latestYear = $null
        $latestMonth = $null
        $latestSnapDate = $null
        $swCount = 0
        $upCount = 0
        foreach ($yd in $yearDirs) {
            $monthDirs = Get-ChildItem -Path $yd.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
            foreach ($md in $monthDirs) {
                $snapFiles = Get-ChildItem -Path $md.FullName -Filter 'snapshot-*.json' -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending
                if ($snapFiles.Count -gt 0) {
                    try {
                        $snap = Get-Content -Path $snapFiles[0].FullName -Raw | ConvertFrom-Json
                        $swCount = $snap.Software.Count
                        $upCount = $snap.Updates.Count
                        $latestSnapDate = $snap.Date
                    } catch { }
                    $latestYear = $yd.Name
                    $latestMonth = $md.Name
                    break
                }
            }
            if ($latestYear) { break }
        }
        $entries += [PSCustomObject]@{
            Computer   = $compFolder
            SnapDate   = if ($latestSnapDate) { $latestSnapDate } else { 'Unknown' }
            SwCount    = $swCount
            UpCount    = $upCount
            Year       = $latestYear
            Month      = $latestMonth
        }
    }

    $entries = $entries | Sort-Object Computer
    $totalComputers = $entries.Count

    $rowsSb = New-Object System.Text.StringBuilder
    foreach ($e in $entries) {
        if ($e.Year -and $e.Month) {
            $link = "<a href='./$($e.Computer)/$($e.Year)/$($e.Month)/report.html'>$(ConvertTo-HtmlEncoded $e.Computer)</a>"
        } else {
            $link = ConvertTo-HtmlEncoded $e.Computer
        }
        $null = $rowsSb.Append(@"
<tr><td>$link</td>
    <td>$(ConvertTo-HtmlEncoded $e.SnapDate)</td>
    <td>$($e.SwCount)</td>
    <td>$($e.UpCount)</td></tr>
"@)
    }
    $rows = $rowsSb.ToString()

    $now = Get-Date
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Computers - Global Inventory</title>
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
  a { color: #1a3a5c; }
  a:hover { text-decoration: underline; }
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
    html.theme-auto body a { color: #80b0e0; }
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
  html.dark body a { color: #80b0e0; }
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
<a class="back-link" href="index.html">&larr; Back to archive</a>
<h1>Computers</h1>
<div class="summary">
  <div class="summary-grid">
    <div class="summary-item"><div class="number">$totalComputers</div><div class="label">Total Computers</div></div>
  </div>
  <div class="meta">Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss'))</div>
</div>

<h2 class="section-title">All Hosts</h2>
<input type="text" id="comp-filter" class="search-box" placeholder="Filter computers..." onkeyup="filterTable('comp-filter','comp-table')">
<table id="comp-table"><thead><tr>
  <th onclick="sortTable('comp-table',0)">Hostname</th>
  <th onclick="sortTable('comp-table',1)">Last Snapshot</th>
  <th onclick="sortTable('comp-table',2)">Software</th>
  <th onclick="sortTable('comp-table',3)">Patches</th>
</tr></thead><tbody>$rows</tbody></table>
</body></html>
"@

    $outputFile = Join-Path $OutputDir "computers.html"
    $html | Out-File -FilePath $outputFile -Encoding utf8
    Write-Host "  Computers page saved: $outputFile"
}

# ---------------------------------------------------------------
# New-WebsiteIndexHtml
# Generates the root Output\index.html with year/month navigation
# linking to combined month index pages across all computers.
# ---------------------------------------------------------------
function New-WebsiteIndexHtml {
    param(
        [string]$OutputDir,
        [string]$HistoryRoot = '',
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

    # Also add months from combined index files (backfilled or report-only months)
    $monthIndexes = Get-ChildItem -Path $OutputDir -Recurse -Filter 'index.html' -ErrorAction SilentlyContinue |
        Where-Object { $_.Directory.Parent.Name -match '^\d{4}$' } |
        Sort-Object FullName
    Write-Host "  Found $($monthIndexes.Count) month index.html files"
    foreach ($mi in $monthIndexes) { Write-Host "    - $($mi.FullName)" }
    Write-Host "  Years before merge: $($years.Keys -join ', ')"
    foreach ($r in $monthIndexes) {
        $rel = $r.Directory.FullName.Substring($OutputDirNorm.Length + 1)
        $parts = $rel -split '[/\\]'
        if ($parts.Count -ge 2) {
            $y = $parts[0]
            $m = $parts[1]
            if (-not $years[$y]) { $years[$y] = @{} }
            if (-not $years[$y][$m]) { $years[$y][$m] = @() }
            if ($years[$y][$m].Count -eq 0) {
                $years[$y][$m] += @{ Computer = ''; Path = $r.FullName }
            }
        }
    }

    Write-Host "  Years after merge: $($years.Keys -join ', ')"
    $years.Keys | Sort-Object | ForEach-Object { $y = $_; $years[$y].Keys | Sort-Object | ForEach-Object { Write-Host "    $y-$_" } }

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
            $compEntries = $years[$y][$m] | Where-Object { $_.Computer -ne '' }
            $compCount = ($compEntries | ForEach-Object { $_.Computer } | Select-Object -Unique).Count
            if ($compCount -eq 0 -and $HistoryRoot) {
                $compCount = (Get-ChildItem -Path $HistoryRoot -Directory -ErrorAction SilentlyContinue |
                    Where-Object { (Get-ChildItem -Path ([System.IO.Path]::Combine($_.FullName, $y, $m)) -Filter 'snapshot-*.json' -ErrorAction SilentlyContinue).Count -gt 0 }
                ).Count
            }
            $compLabel = if ($compCount -gt 0) { " ($compCount PCs)" } else { '' }
            $navHtml += "<a href='./$y/$m/index.html' class='month-link'>$monthName $y$compLabel</a>`n"
        }
        $navHtml += "</div></div>"
    }

    # Build search index for content-aware search
    Write-Host "  Building search index..."
    $searchIndex = @{}
    foreach ($y in $sortedYears) {
        foreach ($m in ($years[$y].Keys | Sort-Object)) {
            $monthKey = "$y-$m"
            $compEntries = $years[$y][$m] | Where-Object { $_.Computer -ne '' }
            foreach ($ce in $compEntries) {
                $compFolder = $ce.Computer
                if ($HistoryRoot -and (Test-Path $HistoryRoot)) {
                    $snapDir = [System.IO.Path]::Combine($HistoryRoot, $compFolder, $y, $m)
                    if (Test-Path $snapDir) {
                        $snapFile = Get-ChildItem -Path $snapDir -Filter 'snapshot-*.json' -ErrorAction SilentlyContinue |
                            Sort-Object LastWriteTime -Descending | Select-Object -First 1
                        if ($snapFile) {
                            try {
                                $snap = Get-Content -Path $snapFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                                foreach ($sw in $snap.Software) {
                                    if (-not $sw.Name) { continue }
                                    $name = ($sw.Name -replace '\s+', ' ').Trim()
                                    if ($name -eq '') { continue }
                                    if (-not $searchIndex.ContainsKey($name)) { $searchIndex[$name] = @{} }
                                    if (-not $searchIndex[$name].ContainsKey($monthKey)) { $searchIndex[$name][$monthKey] = @() }
                                    if ($searchIndex[$name][$monthKey] -notcontains $compFolder) {
                                        $searchIndex[$name][$monthKey] += $compFolder
                                    }
                                    # Index InstallDate
                                    $date = $sw.InstallDate
                                    if ($date -and $date -ne 'Unknown' -and $date -match '\d{4}-\d{2}-\d{2}') {
                                        if (-not $searchIndex.ContainsKey($date)) { $searchIndex[$date] = @{} }
                                        if (-not $searchIndex[$date].ContainsKey($monthKey)) { $searchIndex[$date][$monthKey] = @() }
                                        if ($searchIndex[$date][$monthKey] -notcontains $compFolder) {
                                            $searchIndex[$date][$monthKey] += $compFolder
                                        }
                                    }
                                }
                                # Index computer name
                                $compName = $snap.Computer
                                if ($compName) {
                                    if (-not $searchIndex.ContainsKey($compName)) { $searchIndex[$compName] = @{} }
                                    if (-not $searchIndex[$compName].ContainsKey($monthKey)) { $searchIndex[$compName][$monthKey] = @() }
                                    if ($searchIndex[$compName][$monthKey] -notcontains $compFolder) {
                                        $searchIndex[$compName][$monthKey] += $compFolder
                                    }
                                }
                            } catch {
                                Write-Host "    Warning: Could not parse $($snapFile.FullName): $_"
                            }
                        }
                    }
                }
            }
        }
    }
    Write-Host "  Search index built: $($searchIndex.Count) entries"
    $searchJson = if ($searchIndex.Count -gt 0) {
        (ConvertTo-Json -InputObject $searchIndex -Compress -Depth 5) -replace '</', '<\/'
    } else {
        '{}'
    }

    $now = Get-Date
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Software / Patch Inventory Archive</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
  h1 { color: #1a3a5c; border-bottom: 2px solid #1a3a5c; padding-bottom: 8px; }
  .search-box { margin-bottom: 15px; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; width: 300px; max-width: 100%; box-sizing: border-box; }
  .search-box:focus { outline: none; border-color: #1a3a5c; box-shadow: 0 0 4px rgba(26,58,92,.3); }
  .year-group { background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,.1); margin-bottom: 15px; padding: 15px; }
  .year-heading { margin: 0 0 10px 0; }
  .year-heading a { color: #1a3a5c; text-decoration: none; }
  .year-heading a:hover { text-decoration: underline; }
  .month-list { display: flex; flex-wrap: wrap; gap: 8px; }
  .month-link, .year-link { display: inline-block; background: #e8f0fe; color: #1a3a5c; padding: 8px 16px; border-radius: 4px; text-decoration: none; font-size: 14px; }
  .month-link:hover, .year-link:hover { background: #d0e0f0; }
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
    html.theme-auto body { background: #1a1a2e; color: #e0e0e0; }
    html.theme-auto body h1 { color: #80b0e0; }
    html.theme-auto body .year-group { background: #16213e; }
    html.theme-auto body .month-link, html.theme-auto body .year-link { background: #0f3460; color: #80b0e0; }
    html.theme-auto body .month-link:hover, html.theme-auto body .year-link:hover { background: #1a4a7a; }
    html.theme-auto body .meta { color: #888; }
    html.theme-auto body .failures-link.green { background: #1b5e20; color: #a5d6a7; }
    html.theme-auto body .failures-link.red { background: #b71c1c; color: #ffcdd2; }
  }
  html.dark body { background: #1a1a2e; color: #e0e0e0; }
  html.dark body h1 { color: #80b0e0; }
  html.dark body .year-group { background: #16213e; }
  html.dark body .month-link, html.dark body .year-link { background: #0f3460; color: #80b0e0; }
  html.dark body .month-link:hover, html.dark body .year-link:hover { background: #1a4a7a; }
  html.dark body .meta { color: #888; }
  html.dark body .failures-link.green { background: #1b5e20; color: #a5d6a7; }
  html.dark body .failures-link.red { background: #b71c1c; color: #ffcdd2; }
  #search-results { display: none; margin: 10px 0; }
  .search-result-item { margin: 8px 0; padding: 8px 12px; background: #fff; border: 1px solid #e0e0e0; border-radius: 4px; }
  html.dark body .search-result-item { background: #0f3460; border-color: #2a3a5e; }
  @media (prefers-color-scheme: dark) {
    html.theme-auto body .search-result-item { background: #0f3460; border-color: #2a3a5e; }
  }
  .search-result-item .sr-software { font-weight: bold; color: #1a3a5c; }
  html.dark body .search-result-item .sr-software { color: #80b0e0; }
  @media (prefers-color-scheme: dark) {
    html.theme-auto body .search-result-item .sr-software { color: #80b0e0; }
  }
  .search-result-item .sr-month { color: #2a5a8c; }
  html.dark body .search-result-item .sr-month { color: #4a9eff; }
  @media (prefers-color-scheme: dark) {
    html.theme-auto body .search-result-item .sr-month { color: #4a9eff; }
  }
  .search-result-item .sr-computers { color: #666; }
  html.dark body .search-result-item .sr-computers { color: #aaa; }
  @media (prefers-color-scheme: dark) {
    html.theme-auto body .search-result-item .sr-computers { color: #aaa; }
  }
  .search-results-none { padding: 10px; color: #888; }
</style>
<script>
var searchIndex = $searchJson ;
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
function filterLinks() {
  var input = document.getElementById('search');
  var filter = input.value.trim().toLowerCase();
  var resultsDiv = document.getElementById('search-results');
  var monthLinks = document.querySelectorAll('.month-link');
  if (filter === '') {
    resultsDiv.innerHTML = '';
    resultsDiv.style.display = 'none';
    for (var i = 0; i < monthLinks.length; i++) { monthLinks[i].style.display = ''; }
    return;
  }
  var matchedMonths = {}, matches = [];
  for (var key in searchIndex) {
    if (key.toLowerCase().indexOf(filter) > -1) {
      var monthData = searchIndex[key];
      for (var monthKey in monthData) {
        matchedMonths[monthKey] = true;
        matches.push({ name: key, month: monthKey, computers: monthData[monthKey] });
      }
    }
  }
  if (matches.length > 0) {
    var html = '<div class="search-results-header">' + matches.length + ' result' + (matches.length > 1 ? 's' : '') + ' for "' + escHtml(filter) + '"</div>';
    var seen = {};
    for (var i = 0; i < matches.length; i++) {
      var m = matches[i], uid = m.name + '|' + m.month;
      if (seen[uid]) continue;
      seen[uid] = true;
      var parts = m.month.split('-');
      var monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      var monthName = monthNames[parseInt(parts[1],10)-1] || parts[1];
      html += '<div class="search-result-item"><span class="sr-software">' + escHtml(m.name) + '</span><br><a href="./' + parts[0] + '/' + parts[1] + '/index.html" class="sr-month">' + monthName + ' ' + parts[0] + '</a> <span class="sr-computers">(' + escHtml(m.computers.join(', ')) + ')</span></div>';
    }
    resultsDiv.innerHTML = html;
    resultsDiv.style.display = 'block';
    for (var i = 0; i < monthLinks.length; i++) {
      var link = monthLinks[i], show = false;
      for (var mk in matchedMonths) { if (link.href.indexOf(mk) > -1) { show = true; break; } }
      link.style.display = show ? '' : 'none';
    }
  } else {
    var anyMatch = false;
    for (var i = 0; i < monthLinks.length; i++) {
      var text = monthLinks[i].textContent.toLowerCase();
      var match = text.indexOf(filter) > -1;
      monthLinks[i].style.display = match ? '' : 'none';
      if (match) anyMatch = true;
    }
    resultsDiv.innerHTML = anyMatch ? '<div class="search-results-none">No software/computer matches for "' + escHtml(filter) + '"</div>' : '<div class="search-results-none">No results for "' + escHtml(filter) + '"</div>';
    resultsDiv.style.display = 'block';
  }
}
function escHtml(str) { return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
</script>
</head>
<body>
<button class="theme-toggle" onclick="toggleTheme()">&#9681; Theme</button>
<h1>Software / Patch Inventory Archive</h1>

<input type="text" id="search" class="search-box" placeholder="Search archive..." onkeyup="filterLinks()">
<div id="search-results"></div>

<div class="year-group">
  <h2 style="margin:0 0 10px 0;">Jump to Year</h2>
  <div class="month-list">
$(
    ($sortedYears | ForEach-Object { "<a href='#year-$_' class='year-link'>$_</a>" }) -join "`n"
)
  </div>
  <div style="margin-top: 15px; border-top: 1px solid #e0e0e0; padding-top: 12px;">
    <a href="all-software.html" class="failures-link" style="background:#e8f0fe;color:#1a3a5c;">All Software &amp; Patches</a>
    <a href="computers.html" class="failures-link" style="background:#e8f0fe;color:#1a3a5c;">Computers</a>
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
    html.theme-auto body { background: #1a1a2e; color: #e0e0e0; }
    html.theme-auto body h1 { color: #81c784; }
    html.theme-auto body .success { background: #1b5e20; color: #a5d6a7; }
    html.theme-auto body a.back-link { color: #81c784; }
    html.theme-auto body .meta { color: #a0c0e0; }
    html.theme-auto body .theme-toggle { border-color: #81c784; color: #81c784; }
    html.theme-auto body .theme-toggle:hover { background: #81c784; color: #1a1a2e; }
  }
  html.dark body { background: #1a1a2e; color: #e0e0e0; }
  html.dark body h1 { color: #81c784; }
  html.dark body .success { background: #1b5e20; color: #a5d6a7; }
  html.dark body a.back-link { color: #81c784; }
  html.dark body .meta { color: #a0c0e0; }
  html.dark body .theme-toggle { border-color: #81c784; color: #81c784; }
  html.dark body .theme-toggle:hover { background: #81c784; color: #1a1a2e; }
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
     html.theme-auto body { background: #1a1a2e; color: #e0e0e0; }
     html.theme-auto body h1 { color: #e57373; }
     html.theme-auto body .summary { background: #16213e; }
     html.theme-auto body table { background: #16213e; }
     html.theme-auto body th { background: #b71c1c; }
     html.theme-auto body th:hover { background: #c62828; }
     html.theme-auto body td { border-bottom: 1px solid #2a3a5e; }
     html.theme-auto body tr:hover td { background: #1a2a4e; }
     html.theme-auto body .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
     html.theme-auto body .theme-toggle { border-color: #ef9a9a; color: #ef9a9a; }
     html.theme-auto body .theme-toggle:hover { background: #ef9a9a; color: #1a1a2e; }
   }
   html.dark body { background: #1a1a2e; color: #e0e0e0; }
   html.dark body h1 { color: #e57373; }
   html.dark body .summary { background: #16213e; }
   html.dark body table { background: #16213e; }
   html.dark body th { background: #b71c1c; }
   html.dark body th:hover { background: #c62828; }
   html.dark body td { border-bottom: 1px solid #2a3a5e; }
   html.dark body tr:hover td { background: #1a2a4e; }
   html.dark body .search-box { background: #16213e; border-color: #2a3a5e; color: #e0e0e0; }
   html.dark body .theme-toggle { border-color: #ef9a9a; color: #ef9a9a; }
   html.dark body .theme-toggle:hover { background: #ef9a9a; color: #1a1a2e; }
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
# Backfill-HistoryMonths
# Creates snapshot files for missing months in the current year
# by reconstructing software/patch presence from InstallDate.
# ---------------------------------------------------------------
function Backfill-HistoryMonths {
    param(
        [string]$HistoryRoot,
        [string]$Year
    )

    Write-Host "  Backfill: HistoryRoot=$HistoryRoot"
    $computerDirs = Get-ChildItem -Path $HistoryRoot -Directory -ErrorAction SilentlyContinue
    if ($computerDirs.Count -eq 0) {
        Write-Host "  Backfill: no computer directories found in $HistoryRoot"
        return @()
    }

    $backfilled = @()

    foreach ($compDir in $computerDirs) {
        $compFolder = $compDir.Name

        $snapFiles = Get-ChildItem -Path $compDir.FullName -Recurse -Filter 'snapshot-*.json' -ErrorAction SilentlyContinue
        if ($snapFiles.Count -eq 0) {
            Write-Host "  Backfill: no snapshot files for $compFolder in $($compDir.FullName)"
            continue
        }

        $allSoftware = @{}
        $allPatches = @{}
        $compName = $compFolder

        foreach ($file in $snapFiles) {
            try {
                $snap = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
                $compName = $snap.Computer
                foreach ($sw in $snap.Software) {
                    $key = ($sw.Name -replace '\s+', ' ').Trim().ToLower()
                    if (-not $allSoftware.ContainsKey($key) -or $sw.Version -ne 'Unknown') {
                        $allSoftware[$key] = $sw
                    }
                }
                foreach ($up in $snap.Updates) {
                    $key = ($up.Title -replace '\s+', ' ').Trim().ToLower()
                    if (-not $allPatches.ContainsKey($key)) {
                        $allPatches[$key] = $up
                    } else {
                        $existing = $allPatches[$key]
                        $ed = if ($existing.InstallDate -eq 'Unknown') { '0000-00-00' } else { $existing.InstallDate }
                        $nd = if ($up.InstallDate -eq 'Unknown') { '0000-00-00' } else { $up.InstallDate }
                        if ($nd -gt $ed) { $allPatches[$key] = $up }
                    }
                }
            } catch {
                # skip corrupt
            }
        }

        if ($allSoftware.Count -eq 0 -and $allPatches.Count -eq 0) {
            Write-Host "  Backfill: $compFolder collected no data from $($snapFiles.Count) snapshots"
            continue
        }
        Write-Host "  Backfill: $compFolder collected $($allSoftware.Count) sw, $($allPatches.Count) patches from $($snapFiles.Count) snapshots"

        # Discover unique months from InstallDates in the data
        $foundMonths = @{}
        foreach ($sw in $allSoftware.Values) {
            try { $dt = [datetime]$sw.InstallDate; if ($dt.Year -eq [int]$Year) { $foundMonths[$dt.Month.ToString('00')] = $true } } catch { }
        }
        foreach ($up in $allPatches.Values) {
            try { $dt = [datetime]$up.InstallDate; if ($dt.Year -eq [int]$Year) { $foundMonths[$dt.Month.ToString('00')] = $true } } catch { }
        }

        Write-Host "  Backfill: $compFolder discovered months: $(($foundMonths.Keys | Sort-Object) -join ', ')"

        foreach ($monthStr in ($foundMonths.Keys | Sort-Object)) {
            $monthDir = [System.IO.Path]::Combine($HistoryRoot, $compFolder, $Year, $monthStr)

            $existing = Get-ChildItem -Path $monthDir -Filter 'snapshot-*.json' -ErrorAction SilentlyContinue
            if ($existing.Count -gt 0) {
                Write-Host "      $compFolder ${Year}-${monthStr}: existing snapshot, skipping"
                continue
            }

            $monthSw = $allSoftware.Values | Where-Object {
                try { $dt = [datetime]$_.InstallDate; $dt.Year -eq [int]$Year -and $dt.Month -eq [int]$monthStr }
                catch { $true }
            }

            $monthPatches = $allPatches.Values | Where-Object {
                try { $dt = [datetime]$_.InstallDate; $dt.Year -eq [int]$Year -and $dt.Month -eq [int]$monthStr }
                catch { $false }
            }

            Write-Host "      $compFolder ${Year}-${monthStr}: $($monthSw.Count) sw, $($monthPatches.Count) patches"
            if ($monthSw.Count -eq 0 -and $monthPatches.Count -eq 0) {
                Write-Host "        -> skipping (empty)"
                continue
            }

            Save-HistorySnapshot -Computer $compName -Software $monthSw -Updates $monthPatches `
                -HistoryRoot $HistoryRoot -TargetYear $Year -TargetMonth $monthStr | Out-Null
            Write-Host "      $compFolder ${Year}-${monthStr}: snapshot saved"

            $backfilled += @{ Year = $Year; Month = $monthStr }
        }
    }

    $backfilled
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
Write-Host "Backfilling missing history months..."
$backfilled = Backfill-HistoryMonths -HistoryRoot $HistoryPath -Year $currentYear
if ($backfilled.Count -gt 0) {
    $backfilledMonths = $backfilled | ForEach-Object { "$($_.Year)-$($_.Month)" } | Sort-Object -Unique
    Write-Host "  Backfilled $($backfilled.Count) month(s): $($backfilledMonths -join ', ')"
}
Write-Host "Generating combined month reports..."
$monthsToGenerate = @{}
$historyComps = Get-ChildItem -Path $HistoryPath -Directory -ErrorAction SilentlyContinue
foreach ($compDir in $historyComps) {
    $yearDirs = Get-ChildItem -Path $compDir.FullName -Directory -ErrorAction SilentlyContinue
    foreach ($yearDir in $yearDirs) {
        $monthDirs = Get-ChildItem -Path $yearDir.FullName -Directory -ErrorAction SilentlyContinue
        foreach ($monthDir in $monthDirs) {
            if ((Get-ChildItem -Path $monthDir.FullName -Filter 'snapshot-*.json' -ErrorAction SilentlyContinue).Count -gt 0) {
                $monthsToGenerate["$($yearDir.Name)-$($monthDir.Name)"] = @{ Year = $yearDir.Name; Month = $monthDir.Name }
            }
        }
    }
}
foreach ($entry in ($monthsToGenerate.Values | Sort-Object Year, Month)) {
    Write-Host "  Generating month report for $($entry.Year)-$($entry.Month)..."
    New-MonthReportHtml -Year $entry.Year -Month $entry.Month -OutputDir $OutputPath -HistoryRoot $HistoryPath
}
Write-Host "Generating failures page..."
New-FailuresHtml -Failures $failures -OutputDir $OutputPath
Write-Host "Generating All Software page..."
New-AllSoftwareHtml -HistoryRoot $HistoryPath -OutputDir $OutputPath
Write-Host "Generating Computers page..."
New-ComputersHtml -HistoryRoot $HistoryPath -OutputDir $OutputPath
Write-Host "Generating root website index..."
New-WebsiteIndexHtml -OutputDir $OutputPath -HistoryRoot $HistoryPath -FailureCount $failures.Count

Write-Host "Done."
