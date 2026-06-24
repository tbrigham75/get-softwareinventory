# Software Inventory & Update Reporting

Collects installed software and Windows updates from local/remote machines via registry (Uninstall keys) and the WUA COM API, saves monthly JSON snapshots, and generates browsable HTML reports with year/month navigation, change tracking, and sortable tables.

## Requirements

- Windows PowerShell 5.1+ (the script uses only PS 5.1-compatible APIs)
- **Local:** Run as Administrator
- **Remote:** Admin rights, WinRM enabled, firewall access for Remote Registry (RPC) and WMI/DCOM
- No external modules or NuGet packages required

## Usage

```
.\Get-SoftwareInventoryV7.ps1 [-ComputerName "host1","host2"] [-OutputPath <path>] [-HistoryPath <path>] [-PassThru]
```

| Parameter     | Description |
|---------------|-------------|
| `-ComputerName` | Overrides `hostnames.txt` with explicit list |
| `-OutputPath`   | Root for HTML reports (default: `.\Output`) |
| `-HistoryPath`  | Root for JSON snapshots (default: `.\History`) |
| `-PassThru`     | Return inventory objects to the pipeline |

By default the script reads computer names from `hostnames.txt` (one per line).

## Output Structure

```
Output\
├── index.html              Root archive — year/month navigation + failures link
├── failures.html           Failed computers (separate from inventory data)
└── <Computer>\
    └── YYYY\
        └── MM\
            ├── report.html Per-computer detail: software, updates, changes
            └── snapshot-YYYYMMDD-HHmm.json  Raw persisted data

History\<Computer>\YYYY\MM\snapshot-*.json    Full snapshot history (one per run)

Output\YYYY\MM\index.html                     Combined month view — all hosts together
```

## Key Design Decisions

- **Registry Uninstall keys** (`HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`) — does NOT use `Win32_Product`, which triggers MSI self-repair
- **WUA COM API** — captures all update types (drivers, .NET, definition updates) and works offline
- **Per-computer history** — each host gets its own snapshot folder; change detection compares against the most recent prior run at any time
- **Failures isolation** — failed hosts don't appear in main inventory views; a green/red link on the root index shows failure status
- **Sortable tables** — vanilla JavaScript, no dependencies; click any column header to sort
- **PS 5.1 limitations honored** — no `Join-Path` with >2 segments, no `-WarningVariable`/`4>$null` in nested functions

## Limitations

- Only 64-bit registry view is checked; 32-bit software on 64-bit OS (Wow6432Node) is not enumerated
- Remote registry requires the "Remote Registry" service to be running on targets
- Updates are gathered via the local WUA agent; the target machine's WSUS/SUS settings determine which updates are reported as installed

## Files

| File | Purpose |
|------|---------|
| `Get-SoftwareInventoryV7.ps1` | Main script (active version) |
| `hostnames.txt` | Default list of computers to inventory |
| `History\` | Archived JSON snapshots (auto-created) |
| `Output\` | Generated HTML reports (auto-created) |
