# Phase 4B Windows acceptance

Recorded on the self-hosted Windows worker against tip of
`feature/agent-phase-4b-safe-explorer` after the second re-review fixes.

- Base SHA: `2c6336042127a9553bf974f368f9dbd5af658fe1` (`origin/main`, Phase 4A merge #140)
- Fixture: `tools/DesktopAutomationExplorerFixture` (**DA Explorer Fixture**)
- Evidence folder: `artifacts/phase4b-acceptance-rereview2/` (local; not required in git)

## Exit-gate runs (this tip)

| Run | explorationId | Result |
|---|---|---|
| Dry-run | `20261111T005403384Z-2d15c59a` | PASS — 35 candidates; Exit/Reset/Print/Archive/Delete/Submit/Save/Purchase absent |
| Bounded navigate | `20261111T005404546Z-08e388bb` | PASS — stopReason=`max-actions-budget`, actions=40, states=5 |
| Modal/second + backtrack | `20261111T005311044Z-b7703808` | PASS — `newWindow` for **DA Explorer Second**; backtrack dismiss closed second windows; dialog/second invoke via explicit allow |

## Exit gate checklist

| Check | Evidence |
|---|---|
| Complete fixture traversal | Dry-run candidates cover expand/tab/menuOpen/scroll/focus/select/invoke-allows; navigate exercised transitions |
| Modal / second-window | Run `…b7703808` — `newWindow\|ExplorerOpenSecondButton\|New window(s): DA Explorer Second` (dialogs opened via invoke allow; fixture uses owned non-blocking dialog so dumpuia can stabilize) |
| External-window rejection | `listwindows` includes `processName`; new windows missing/`notepad` fail closed in unit + detector tests |
| Process identity fail-closed | Toolhelp32 parent PID; pin root for run; probe failure rejects `ValidateNewWindow` |
| Budget maxStates/maxActions/maxDepth/duration | Navigate `max-actions-budget`; unit tests for maxDepth=0 and maxDepth=1 |
| Cancellation + partial artifacts | Unit: cancelled run writes summary; live runs write manifest/states/transitions/summary |
| Stabilization timeout | Unit coverage + live stabilize quiet period honored |
| Backtracking | `allowCloseChildWindowsForBacktrack=true`; stack push from `AddedWindows`; switch+dismiss then switch-parent |
| Stale refs | Live navigate records `staleRef` when SnapshotId/fingerprint drift; NodeRef only when SnapshotId matches |
| Fingerprint interaction | dumpuia interaction → normalizer `ComputeHash` + `StateFingerprinter` |
| Schema enforcement | Runtime `exploration-policy.schema.json`; misspelled `maxActons` rejected |

## Agent suite

`Phase4bSafeExplorerTests`: **27** passed locally before push.
