# Dark Mode & Consistency Fixes

## Medium Priority

### Fix 1: Remote Backfill-HistoryMonths old normalization
**File:** `Get-RemoteInventory.ps1` lines 2332, 2338
**Current:** `($sw.Name -replace '\s+', ' ').Trim().ToLower()` and `($up.Title -replace '\s+', ' ').Trim().ToLower()`
**Fix:** `Normalize-SoftwareName $sw.Name` and `Normalize-SoftwareName $up.Title`

### Fix 2: Remote month view summary links missing anchors
**File:** `Get-RemoteInventory.ps1` lines 1114-1115
**Current:** `href="../../../all-software.html"`
**Fix:** `href="../../../all-software.html#software-section"` and `href="../../../all-software.html#patches-section"`

### Fix 3: Index all-software.html dark mode missing .summary-item .label
**File:** `Get-InventoryIndex.ps1` — in `New-AllSoftwareIndexHtml` dark mode blocks (~lines 753-778)
**Fix:** Add `.summary-item .label { color: #a0c0e0; }` to both theme-auto and html.dark blocks

### Fix 4: Remote all-software .label color inconsistency
**File:** `Get-RemoteInventory.ps1` lines 1313, 1327
**Current:** `color: #aab`
**Fix:** `color: #a0c0e0` (matching all other pages)

### Fix 5: Remote all-software dark mode missing th:hover
**File:** `Get-RemoteInventory.ps1` — in `New-AllSoftwareHtml` dark mode blocks
**Fix:** Add `th:hover { background: #1a4a7a; }` to both theme-auto and html.dark blocks

### Fix 6: Systemic dark mode — a.back-link and .theme-toggle
**All affected pages across all 3 scripts need these rules added to both theme-auto and html.dark blocks:**
- `a.back-link { color: #80b0e0; }` / `.theme-toggle { border-color: #80b0e0; color: #80b0e0; }` / `.theme-toggle:hover { background: #80b0e0; color: #1a1a2e; }`

**Remote pages affected:**
- `New-InventoryHtmlReport` (~line 710-716)
- `New-MonthReportHtml` (~line 1050-1060)
- `New-AllSoftwareHtml` (~line 1318-1331)
- `New-ComputersHtml` (~line 1585-1606)
- `New-FailuresHtml` (with failures, ~line 2210)

**Index pages affected:**
- `New-LocalHtmlReport` (~line 300-320)
- `New-AllSoftwareIndexHtml` (~line 760-778)
- Month view Step 4 (~line 1200-1218)
- `New-WebsiteIndexHtml` (~line 1430-1450)

**Local pages affected:**
- `New-HtmlReport` (~line 480-494)

## Low Priority

### Fix 7: Remote all-software td border color
**File:** `Get-RemoteInventory.ps1` lines 1316, 1330
**Current:** `#2a2a4e`
**Fix:** `#2a3a5e`

### Fix 8: Remote back-link text
**File:** `Get-RemoteInventory.ps1` — all pages
**Current:** "Back to archive"
**Fix:** "Back to inventory" (matching Local/Index) — OR leave as-is if "archive" is the intended Remote terminology

### Fix 9: Remote Backfill-HistoryMonths snapshot sort
**File:** `Get-RemoteInventory.ps1` line 2317
**Fix:** Add `| Sort-Object LastWriteTime` to `Get-ChildItem`

### Fix 10: Remote root index search month-link filtering
**File:** `Get-RemoteInventory.ps1` line 2000
**Current:** `link.href.indexOf(mk)` where mk="2024-01" but href contains "2024/01"
**Fix:** Use `link.getAttribute('href').indexOf(parts[0] + '/' + parts[1])` or similar
