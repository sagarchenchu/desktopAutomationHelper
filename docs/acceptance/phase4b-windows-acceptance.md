# Phase 4B Windows acceptance

Recorded on the self-hosted Windows worker against tip of
`feature/agent-phase-4b-safe-explorer` after the **third** re-review fixes.

- Base SHA: `2c6336042127a9553bf974f368f9dbd5af658fe1` (`origin/main`, Phase 4A merge #140)
- Fixture: `tools/DesktopAutomationExplorerFixture` (**DA Explorer Fixture**)
- Evidence folder: `artifacts/phase4b-acceptance-rereview3/` (local; not required in git)

## Exit-gate runs (this tip)

| Run | explorationId | Result |
|---|---|---|
| Dry-run | `20261111T233308403Z-56535d4c` | PASS — 32 candidates; Exit/Reset/Print/Archive/Delete/Submit/Save/Purchase absent |
| Bounded navigate | `20261111T233309728Z-d7a8f9ce` | PASS — stopReason=`max-actions-budget`, actions=40, states=4 |
| Modal/second + backtrack | `20261111T233346435Z-1384952e` | PASS — `modalDialog` for **DA Explorer Dialog** (BeginInvoke+ShowDialog); HWND WindowPattern backtrack `closedWindow` |

## Exit gate checklist / coverage matrix

| Check | Evidence |
|---|---|
| Complete fixture traversal | Dry-run candidates cover expand/collapse/tab/menuOpen/scroll/focus/select/invoke-allows |
| True modal | Run `…1384952e` — `modalDialog\|ExplorerOpenDialogButton\|New window(s): DA Explorer Dialog` (fixture `BeginInvoke` + `ShowDialog`) |
| HWND backtrack | Same run — `backtrack:close:0x…` / `closedWindow` / detail contains `WindowPattern`; depth preserved (`1`/`2`) |
| NewWindow ≠ NoOp | Unit: `Coordinator_NewWindow_IndependentOfNoOpFingerprint_SwitchesAndEnqueues`; live modal path switches to verified HWND before enqueue |
| External-window rejection | `listwindows` includes `processName`; new windows missing/`notepad` fail closed in unit + detector tests |
| Pinned process identity | Crawl-scoped validator; root PID+session pin; ancestry must reach pin; ProcessName==root shortcut removed |
| Budget maxStates/maxActions/maxDepth/duration | Navigate `max-actions-budget`; units for exact `maxStates`, `total-duration-budget`, maxDepth 0/1 |
| Cancellation + partial artifacts | Unit: cancelled run writes summary; live runs write manifest/states/transitions/summary |
| Stabilization timeout | Unit: `Coordinator_RecordsStabilizationTimeout_WhenFingerprintNeverQuiets` |
| Stale refs | Live navigate records `staleRef` when SnapshotId/fingerprint drift |
| Schema enforcement | Runtime `exploration-policy.schema.json`; misspelled `maxActons` rejected |

## Agent suite

`Phase4bSafeExplorerTests`: **32** passed locally before push (includes backtrack, NewWindow, stabilization timeout, exact maxStates, total-duration, window-stack HWND requirements).
