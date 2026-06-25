# Software Inventory & Update Reporting

Collects installed software and Windows updates from local/remote machines via registry (Uninstall keys) and the WUA COM API, saves monthly JSON snapshots, and generates browsable HTML reports with year/month navigation, change tracking, deduplication, historical backfill, and sortable tables.

## Requirements

- Windows PowerShell 5.1+ (the script uses only PS 5.1-compatible APIs)
- **Local:** Run as Administrator
- **Remote:** Admin rights, WinRM enabled, firewall access for Remote Registry (RPC) and WMI/DCOM
- No external modules or NuGet packages required

## Usage

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

## Output Structure

```
Output\
├── index.html              Root archive — year/month navigation + failures + "All Software & Patches" link
├── failures.html           Failed computers (separate from inventory data)
├── all-software.html       Deduplicated global view of all software and patches across all history
├── YYYY\
│   └── MM\
│       └── index.html      Combined month view — all hosts together, filtered to that month
└── <Computer>\
    └── YYYY\
        └── MM\
            ├── report.html Per-computer detail: software, updates, changes
            └── snapshot-YYYYMMDD-HHmm.json  Raw persisted data

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

## Key Design Decisions

- **Registry Uninstall keys** (`HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`) — does NOT use `Win32_Product`, which triggers MSI self-repair
- **WUA COM API** — captures all update types (drivers, .NET, definition updates) and works offline
- **Per-computer history** — each host gets its own snapshot folder; change detection compares against the most recent prior run at any time
- **Failures isolation** — failed hosts don't appear in main inventory views; a green/red link on the root index shows failure status
- **Sortable tables** — vanilla JavaScript, no dependencies; click any column header to sort
- **PS 5.1 limitations honored** — no `Join-Path` with >2 segments (uses `[System.IO.Path]::Combine` instead), no `-WarningVariable`/`4>$null` in nested functions
- **`[datetime]` cast for date handling** — robust against `ConvertFrom-Json` behavior differences between PS 5.1 (strings) and PS 7+ (DateTime objects)
- **Data-driven month discovery** — `Backfill-HistoryMonths` discovers months from actual snapshot InstallDates rather than iterating month numbers, naturally handling non-contiguous dates

## Limitations

- 32-bit software on 64-bit OS (Wow6432Node) is enumerated; however, per-user (HKCU) software is only collected locally, not via remote registry fallback
- Remote registry requires the "Remote Registry" service to be running on targets
- Updates are gathered via the local WUA agent; the target machine's WSUS/SUS settings determine which updates are reported as installed

## Files

| File | Purpose |
|------|---------|
| `Get-SoftwareInventory.ps1` | Main script (~2065 lines, ~25 functions) |
| `hostnames.txt` | Default list of computers to inventory |
| `History\` | Archived JSON snapshots (auto-created) |
| `Output\` | Generated HTML reports (auto-created) |
