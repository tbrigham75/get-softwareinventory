# Software Inventory & Update Reporting

Collects installed software and Windows updates from local/remote machines via registry (Uninstall keys) and the WUA COM API, saves monthly JSON snapshots, and generates browsable HTML reports with year/month navigation, change tracking, deduplication, historical backfill, and sortable tables.

This repository contains three scripts that serve different environments:

| Script | Environment | Remote Access |
|--------|------------|---------------|
| `Get-SoftwareInventory.ps1` | WinRM-enabled networks | Uses WinRM/remote registry for remote collection |
| `Get-LocalInventory.ps1` | Locked-down networks | No remote access — runs locally via Scheduled Task |
| `Get-InventoryIndex.ps1` | Central reporting | No remote access — reads files from a network share |

## Requirements

- Windows PowerShell 5.1+ (all scripts use only PS 5.1-compatible APIs)
- No external modules or NuGet packages required

Per script:

| Script | Requirements |
|--------|-------------|
| `Get-SoftwareInventory.ps1` | Local: Administrator. Remote: Admin rights, WinRM enabled, firewall access for Remote Registry (RPC) and WMI/DCOM |
| `Get-LocalInventory.ps1` | Local admin rights on each machine; write access to the inventory share |
| `Get-InventoryIndex.ps1` | Read access to the inventory share; write access to generate `index.html` |

---

## Get-SoftwareInventory.ps1 — WinRM-Based Collection

The original script for WinRM-enabled environments. Collects inventory from one or more machines (local + remote), saves JSON snapshots, and generates a full HTML report website.

### Usage

```
.\Get-SoftwareInventory.ps1 [-ComputerName "host1","host2"] [-OutputPath <path>] [-HistoryPath <path>] [-PassThru] [-ExportCsv] [-ThrottleLimit <n>]
```

| Parameter | Description |
|-----------|-------------|
| `-ComputerName` | Overrides `hostnames.txt` with explicit list |
| `-OutputPath` | Root for HTML reports (default: `.\Output`) |
| `-HistoryPath` | Root for JSON snapshots (default: `.\History`) |
| `-PassThru` | Return inventory objects to the pipeline |
| `-ExportCsv` | Export software and updates to CSV files alongside HTML |
| `-ThrottleLimit` | Max concurrent remote hosts during parallel collection (default: 5, max: 50) |

By default the script reads computer names from `hostnames.txt` (one per line).

### Output Structure

```
Output\
├── index.html              Root archive — year/month navigation + search index + page links
├── failures.html           Failed computers (separate from inventory data)
├── all-software.html       Deduplicated global view of all software and patches across all history
├── computers.html          List of all unique hostnames with per-computer links
├── hosts\
│   └── <Computer>\
│       └── YYYY\
│           └── MM\
│               ├── report.html      Per-computer inventory report
│               └── software-*.csv   (optional) CSV exports
└── years\
    └── YYYY\
        └── MM\
            └── index.html  Combined month view — all hosts filtered to that month

History\<Computer>\YYYY\MM\snapshot-*.json    Full snapshot history (one per run)
```

## Features

### Deduplication
- **`Merge-UpdateDuplicates`** — Windows Update entries are deduplicated by normalized Title, keeping the latest InstallDate per group. Applied at collection time for both local and remote hosts.
- Software is deduplicated by normalized Name (version-agnostic grouping) in the All Software & Patches view.

### Month-Filtered Views
- The combined month report (`Output/YYYY/MM/index.html`) shows only software and patches whose InstallDate falls within that specific year/month — each month shows "activity" (items installed during that month).
- Per-computer reports show the full inventory snapshot for that computer/month.

### All Software & Patches
- A global page (`Output/all-software.html`) that scans all historical snapshots across every computer/month.
- Two searchable/sortable tables: **3rd Party Software** (deduped by name, keeps latest Version) and **Windows Patches** (deduped by Title, keeps latest InstallDate).
- Summary grid with deduplicated unique counts.

### Historical Backfill
- **`Backfill-HistoryMonths`** discovers unique (Year, Month) pairs from actual InstallDate values in snapshot data and reconstructs missing monthly snapshots.
- One-time backfill per month — existing snapshots are skipped.
- All discovered history months get a combined month report generated regardless of how the snapshot was created.

### Parallel Remote Collection
- Remote hosts are inventoried concurrently via a single `Invoke-Command` call (throttled by `-ThrottleLimit`); local hosts are always processed sequentially.
- Fallback from WinRM to remote registry if WMI/WinRM calls fail.

### Navigation PC Count
- Month buttons in the root index show the number of contributing PCs, falling back to scanning History files if no report entries exist for that month.

### Search Index
- The root index includes a client-side search index built from all snapshot data — searches software names, versions, update titles, install dates, and computer names.
- Search results link to the corresponding combined month view, filtered by matching computers.

### Historical Per-Computer Reports
- When generating the website, the script backfills per-computer HTML reports for every historical month found in the snapshot archive, so hostname links in month reports always work.

## Key Design Decisions

- **Registry Uninstall keys** (`HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`) — does NOT use `Win32_Product`, which triggers MSI self-repair
- **WUA COM API** — captures all update types (drivers, .NET, definition updates) and works offline
- **Per-computer history** — each host gets its own snapshot folder; change detection compares against the most recent prior run at any time
- **Failures isolation** — failed hosts don't appear in main inventory views; a green/red link on the root index shows failure status
- **Sortable tables** — vanilla JavaScript, no dependencies; click any column header to sort
- **Dark/light theme** — vanilla JavaScript, propagates across all pages via `?theme=` URL query parameters with `localStorage`/`matchMedia` fallback
- **PS 5.1 limitations honored** — no `Join-Path` with >2 segments (uses `[System.IO.Path]::Combine` instead), no `-WarningVariable`/`4>$null` in nested functions
- **`[datetime]` cast for date handling** — robust against `ConvertFrom-Json` behavior differences between PS 5.1 (strings) and PS 7+ (DateTime objects)
- **Data-driven month discovery** — `Backfill-HistoryMonths` discovers months from actual snapshot InstallDates rather than iterating month numbers, naturally handling non-contiguous dates

## Limitations (Get-SoftwareInventory.ps1)

- 32-bit software on 64-bit OS (Wow6432Node) is enumerated; however, per-user (HKCU) software is only collected locally, not via remote registry fallback
- Remote registry requires the "Remote Registry" service to be running on targets
- Updates are gathered via the local WUA agent; the target machine's WSUS/SUS settings determine which updates are reported as installed

---

## Get-LocalInventory.ps1 — Local Collection to Network Share

For environments where WinRM, remote registry, and other remote management tools are locked down. Runs locally on each machine (e.g. via Scheduled Task) and writes collected data to a shared network folder.

### How It Works

1. Validates the share is accessible and writable (temp file test)
2. Collects installed software from registry (same logic as `Get-SoftwareInventory.ps1`)
3. Collects Windows updates via WUA COM API with hotfix fallback
4. Saves a timestamped JSON snapshot to the share
5. Generates an HTML report for the host on the share
6. All errors are written to a local `errors.log` file

### Usage

```
.\Get-LocalInventory.ps1 [-SharePath "\\server\inventory"] [-ExportCsv] [-LogPath "C:\logs\errors.log"]
```

| Parameter | Description |
|-----------|-------------|
| `-SharePath` | UNC path or drive letter to the inventory share (default: `\\fileserver\inventory`) |
| `-ExportCsv` | Export software and updates to CSV files alongside JSON/HTML |
| `-LogPath` | Full path to the local error log (default: `<script directory>\errors.log`) |

### Share Layout

Each machine writes into a folder named after its hostname, with year/month subfolders:

```
\\fileserver\inventory\
├── index.html                          Generated by Get-InventoryIndex.ps1
├── _logs\                              Script error logs (if share is unavailable)
│   └── <hostname>-errors.log
├── HOSTNAME1\
│   ├── 2026\
│   │   └── 07\
│   │       ├── snapshot-20260715-0830.json
│   │       ├── report.html
│   │       ├── software-20260715-0830.csv   (optional)
│   │       └── updates-20260715-0830.csv    (optional)
│   └── ...
├── HOSTNAME2\
│   └── ...
└── ...
```

### Deployment via Scheduled Task

Register the task on each target machine:

```powershell
$action  = New-ScheduledTaskAction -Execute 'pwsh.exe' `
    -Argument '-NoProfile -NonInteractive -File "\\fileserver\share\Get-LocalInventory.ps1" -SharePath "\\fileserver\inventory"'
$trigger = New-ScheduledTaskTrigger -Daily -At "02:00"
$principal = New-ScheduledTaskPrincipal -UserId "DOMAIN\svc_inventory" -RunLevel Highest
Register-ScheduledTask -TaskName "SoftwareInventory" -Action $action -Trigger $trigger -Principal $principal -Description "Local software inventory collector"
```

The service account needs:
- Local admin rights on each target machine
- Write access to the inventory share

The script uses the UNC path directly — no drive mapping is needed or performed. The service account's security context handles authentication to the file server automatically.

---

## Get-InventoryIndex.ps1 — Central Index Generator

Scans the inventory share, reads each host's latest snapshot, and generates a central `index.html` at the share root with links to every host's report.

### How It Works

1. Scans the share root for hostname folders (skips `_logs`)
2. Reads the latest JSON snapshot from each host folder
3. Generates `index.html` with summary stats and a sortable, filterable table

### Usage

```
.\Get-InventoryIndex.ps1 [-SharePath "\\server\inventory"] [-LogPath "C:\logs\index-errors.log"]
```

| Parameter | Description |
|-----------|-------------|
| `-SharePath` | UNC path or drive letter to the inventory share (default: `\\fileserver\inventory`) |
| `-LogPath` | Full path to the local error log (default: `<script directory>\errors.log`) |

### What the Index Shows

- Summary bar: total hosts, total software entries, total update entries, last scan date
- Sortable/filterable table with columns: Hostname (link to report), Software count, Updates count, Last scan date
- Dark/light theme toggle, client-side search/filter
- Same visual style as the per-host reports (Segoe UI, `#1a3a5c` theme)

### Scheduling

Run after `Get-LocalInventory.ps1` has finished writing to the share:

```powershell
$action  = New-ScheduledTaskAction -Execute 'pwsh.exe' `
    -Argument '-NoProfile -NonInteractive -File "C:\Scripts\Get-InventoryIndex.ps1" -SharePath "\\fileserver\inventory"'
$trigger = New-ScheduledTaskTrigger -Daily -At "04:00"
Register-ScheduledTask -TaskName "InventoryIndex" -Action $action -Trigger $trigger -Description "Inventory index generator"
```

---

## Files

| File | Purpose |
|------|---------|
| `Get-SoftwareInventory.ps1` | WinRM-based collection script (~2628 lines, ~25 functions) |
| `Get-LocalInventory.ps1` | Local collection to network share (694 lines, 12 functions) |
| `Get-InventoryIndex.ps1` | Central index generator for share (405 lines, 8 functions) |
| `hostnames.txt` | Default list of computers to inventory (used by `Get-SoftwareInventory.ps1`) |
| `History\` | Archived JSON snapshots (auto-created by `Get-SoftwareInventory.ps1`) |
| `Output\` | Generated HTML reports (auto-created by `Get-SoftwareInventory.ps1`) |
