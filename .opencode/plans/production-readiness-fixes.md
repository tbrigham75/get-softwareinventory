# Production Readiness Fixes

## MUST FIX (High)

### Fix 1: Index host table hrefs use backslashes
**File:** `Get-InventoryIndex.ps1` line 953
**Current:** `$reportRelPath = $reportFiles[0].FullName.Substring($SharePath.Length + 1)`
**Fix:** `$reportRelPath = ($reportFiles[0].FullName.Substring($SharePath.Length + 1)) -replace '\\', '/'`
**Why:** Produces `inventory\HOSTNAME\2025\07\report.html` in hrefs. Browsers may not handle backslashes.

### Fix 2: Remote per-host clickable summaries link to wrong page
**File:** `Get-RemoteInventory.ps1` lines 793-794
**Current:** Both link to `../../../../all-software.html`
**Fix:** Software links to `#allSw-table`, Patches links to `#allUp-table` (matching Local/Index behavior)
**Why:** Clicking counts navigates away from the host report, losing context.

### Fix 3: Remote "Jump to Year" links broken
**File:** `Get-RemoteInventory.ps1` line 1743
**Current:** `<div class='year-group'>`
**Fix:** `<div class='year-group' id='year-$y'>`
**Why:** Jump-to-year bar at line 2026 targets `#year-$y` but the div has no id attribute.

## SHOULD FIX (Medium)

### Fix 4: Remote all-software page missing h2 CSS rule
**File:** `Get-RemoteInventory.ps1` line 1285
**Current:** No `h2` rule in base styles
**Fix:** Add `h2 { color: #1a3a5c; }` after the h1 rule
**Also:** Add h2 to dark mode selectors at lines 1308 and 1321:
- Line 1308: `html.theme-auto body h1, html.theme-auto body h2 { color: #80b0e0; }`
- Line 1321: `html.dark body h1, html.dark body h2 { color: #80b0e0; }`

### Fix 5: Remote all-software page missing .summary-item .label dark mode
**File:** `Get-RemoteInventory.ps1` after line 1311 and after line 1324
**Fix:** Add to both theme-auto and dark blocks:
- `html.theme-auto body .summary-item .label { color: #aab; }`
- `html.dark body .summary-item .label { color: #aab; }`

### Fix 6: Remote Compare-Snapshots uses old normalization
**File:** `Get-RemoteInventory.ps1` lines 559 and 567
**Current:** `($_.Name -replace '\s+', ' ').Trim().ToLower()`
**Fix:** `Normalize-SoftwareName $_.Name`
**Also lines 582-589** (update normalization): Use `Normalize-SoftwareName $_.Title` (which includes lowercasing)

### Fix 7: Index month view says "Windows Updates"
**File:** `Get-InventoryIndex.ps1` line 1312
**Current:** `Windows Updates`
**Fix:** `Windows Patches` (consistent with every other page)

## Items Reviewed and Confirmed OK

- **filterTable toUpperCase in all-software pages**: Both Index (line 817) and Remote (line 1370) use the same toUpperCase + cell-by-cell approach for all-software pages. Other pages use toLowerCase + tbody. This is a consistent pattern, not a bug.
- **Backfill-HistoryMonths old normalization** (Remote line 2329, Index line 475): Intentionally kept different — these are hashtable key builders for snapshot reconstruction, not dedup. Changing them could alter backfill behavior.
- **Search index normalization** (Remote line 1784): JS search uses `.toLowerCase()` client-side, so results are found regardless.
- **scriptVersion mismatch** (Remote 7.0 vs Index/Local 1.0): These scripts have different version histories. Acceptable.
- **Double hotfix fallback** (Local 656, Remote 350): Redundant but harmless — second call is a no-op if first returned data.
