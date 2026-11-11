# Phase 4B Windows acceptance

Recorded on the self-hosted Windows worker against branch
`feature/agent-phase-4b-safe-explorer`.

- Base SHA: `2c6336042127a9553bf974f368f9dbd5af658fe1` (`origin/main`, Phase 4A merge #140)
- Fixture: `tools/DesktopAutomationExplorerFixture` (**DA Explorer Fixture**)

## Exit gate (must all pass)

| Check | Result / evidence |
|---|---|
| Complete fixture traversal (tabs, tree expand, menu open, list/scroll/focus as policy allows) | PASS — dry-run + bounded navigate; leaf Exit/Reset/Print/Archive never MenuOpen candidates |
| Modal / second-window in-scope handling | PASS — dialog + second window titles allowed; outcomes `ModalDialog` / `NewWindow` recorded |
| External-window rejection | PASS — new window without matching processName / foreign process → `scope-violation` (listwindows includes `processName`) |
| Process identity fail-closed | PASS — missing processName/processId rejected; executablePath/SHA256 enforced when set |
| Budget: maxStates | PASS — stopReason=`max-states-budget`; states never exceed policy max |
| Budget: maxActions | PASS — run evidence stopReason=`max-actions-budget` |
| Budget: maxDepth | PASS — no frontier re-enqueue at maxDepth; depth clamp hard |
| Budget: totalDuration | PASS — deadline CTS cancels in-flight action/stabilization; stopReason=`total-duration-budget` |
| Cancellation + partial artifacts | PASS — cancelled run writes manifest/states/transitions/summary |
| Stabilization timeout | PASS — outcome `StabilizationTimeout` when quiet period never met |
| Backtracking / cycle protection | PASS — FrontierQueue rejects duplicate state/action pairs; no unbounded recursion |
| Stale refs fail closed | PASS — unambiguous identity rematch only; re-authorize live node or skip (`StaleRef`) |
| Fingerprint interaction state | PASS — expand/selected/toggle changes fingerprint (no false NoOp/cycle on tab/expand) |
| Destructive controls refused | PASS — Delete/Submit/Save/Purchase never candidates |
| Dry-run lists permitted candidates only | PASS — no UI mutation |

## Prior recorded runs

| Check | Result |
|---|---|
| Dry-run lists permitted candidates | PASS — explorationId `20261110T225354387Z-f4e240f4` |
| Destructive controls absent from dry-run | PASS — Delete/Submit/Save/Purchase never candidates |
| Bounded navigate respects maxActions | PASS — run `20261110T225522854Z-e285ccd0`, stopReason=`max-actions-budget`, actions=6 |
| No typing / transactional actions | PASS |
| Process stayed on fixture | PASS — processName `DesktopAutomationExplorerFixture` |
| Partial/budget artifacts written | PASS — manifest/states/transitions/summary present |

Evidence under `artifacts/phase4b-acceptance/` (local; not required in git).

Agent suite `Phase4bSafeExplorerTests` covers fake-graph budgets, cancellation, stale-ref,
leaf-menu≠MenuOpen, fingerprint interaction, process-identity fail-closed, and missing
processName external-window rejection independent of the live UI.
