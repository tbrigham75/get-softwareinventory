# Fixed Bugs

## Merge-SoftwareDuplicates missing from remote script blocks
**Date:** 2026-07-14
**File:** Get-SoftwareInventory.ps1
**Severity:** High

`Merge-SoftwareDuplicates` was not included in either the individual WinRM
path or the parallel remote collection script block. This caused
`Get-LocalSoftware` to fail silently on remote hosts, returning 0 software
entries. Updates were unaffected.

**Fix:** Added `Merge-SoftwareDuplicates` to both the individual WinRM
`Invoke-Command` script block and the parallel remote collection script
block using the same `${function:Merge-SoftwareDuplicates}` pattern used
for all other functions.

---

## `.` to `localhost` folder name conversion inconsistency
**Date:** 2026-07-14
**File:** Get-SoftwareInventory.ps1
**Severity:** Medium

`Save-HistorySnapshot` converted computer name `.` to `localhost` for
folder names, but all other functions that computed `$compFolder` from
`$Computer` did not. This caused path mismatches: snapshots saved under
`History/localhost/...` could not be found by `Load-HistorySnapshot`
(searching `History/./...`), breaking change tracking. Report links were
also broken.

**Fix:** Extracted a shared `ConvertTo-SafeFolderName` helper function
and replaced all 9 manual `$compFolder` computations with calls to it.

---

## Local host failures not tracked in $failures
**Date:** 2026-07-14
**File:** Get-SoftwareInventory.ps1
**Severity:** Low

When local software/update collection failed, errors were only printed
to the console via `Write-Warning`. No failure record was added to
`$failures`, so `failures.html` showed zero failures even when local
collection failed.

**Fix:** Added `$failures +=` in both local catch blocks, matching the
pattern used for remote failures.

---

## Silent error swallowing in remote collection
**Date:** 2026-07-14
**File:** Get-SoftwareInventory.ps1
**Severity:** Low

The parallel remote collection script block had empty `catch { }` blocks
that silently swallowed errors. Remote hosts could appear to have 0
software/updates with no error recorded.

**Fix:** Added `Write-Warning` logging in the catch blocks and propagated
error information back in the result hash so it can be added to
`$failures`.
