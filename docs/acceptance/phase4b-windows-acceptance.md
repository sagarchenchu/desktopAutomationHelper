# Phase 4B Windows acceptance

Recorded on the self-hosted Windows worker against tip of
`feature/agent-phase-4b-safe-explorer` after the **fourth** re-review fixes.

- Base SHA: `2c6336042127a9553bf974f368f9dbd5af658fe1` (`origin/main`, Phase 4A merge #140)
- Fixture: `tools/DesktopAutomationExplorerFixture` (**DA Explorer Fixture**)
- **PR-auditable evidence:** [`phase4b-rereview4/`](./phase4b-rereview4/) (summaries, transitions, policies committed in-repo)

## Exit-gate runs (this tip)

| Run | explorationId | Result | Evidence |
|---|---|---|---|
| Dry-run | `20261111T235005558Z-faa3ad86` | PASS — dry-run completed | [summary](./phase4b-rereview4/20261111T235005558Z-faa3ad86/summary.json) / [stdout](./phase4b-rereview4/dry-run.json) |
| Bounded navigate | `20261111T235006862Z-f56059b0` | PASS — `max-actions-budget`, actions=40, states=4 | [summary](./phase4b-rereview4/20261111T235006862Z-f56059b0/summary.json) / [transitions](./phase4b-rereview4/20261111T235006862Z-f56059b0/transitions.jsonl) |
| Modal + HWND backtrack | `20261111T235043623Z-d66161fd` | PASS — `modalDialog`×2 + `closedWindow`×2 via HWND WindowPattern | [summary](./phase4b-rereview4/20261111T235043623Z-d66161fd/summary.json) / [transitions](./phase4b-rereview4/20261111T235043623Z-d66161fd/transitions.jsonl) |

## Exit gate checklist / coverage matrix

| Check | Evidence |
|---|---|
| Complete fixture traversal | Dry-run candidates in [dry-run.json](./phase4b-rereview4/dry-run.json) |
| True modal | `modalDialog` + `New window(s): DA Explorer Dialog` in modal transitions |
| HWND-bound WindowPattern backtrack | `backtrack:close:0x…` / detail contains `WindowPattern`; close args are hwnd+processId (driver method `windowPattern`) |
| Post-switch HWND verify | Coordinator requires snapshot HWND == requested HWND before close/enqueue |
| Window identity = HWND+PID | Unit: title change alone is neither NewWindow nor ClosedWindow |
| Session pin fail-closed | Live session required; snapshot≠live rejected; pinned session enforced for root/child/new-window |
| NewWindow ≠ NoOp | Unit + live modal path switches verified HWND then re-observes |
| Budget / cancel labeling | Empty-frontier path checks budget/cancel before `frontier-exhausted`; failed/timed-out closes count toward `maxActions` |
| External-window rejection | Detector + identity unit coverage |
| Schema enforcement | Runtime schema rejects misspelled budgets |

## Agent suite

`Phase4bSafeExplorerTests`: **34** passed locally before push.
