# Phase 4B Windows acceptance

Recorded on the self-hosted Windows worker against tip of
`feature/agent-phase-4b-safe-explorer` after the **fifth** re-review fixes.

- Base SHA: `2c6336042127a9553bf974f368f9dbd5af658fe1` (`origin/main`, Phase 4A merge #140)
- Fixture: `tools/DesktopAutomationExplorerFixture` (**DA Explorer Fixture**)
- **PR-auditable evidence:** [`phase4b-rereview5/`](./phase4b-rereview5/) (summaries, transitions, policies committed in-repo)

## Exit-gate runs (this tip)

| Run | explorationId | Result | Evidence |
|---|---|---|---|
| Dry-run | `20261112T024749897Z-09fc64db` | PASS — dry-run completed | [summary](./phase4b-rereview5/20261112T024749897Z-09fc64db/summary.json) / [stdout](./phase4b-rereview5/dry-run.json) |
| Bounded navigate | `20261112T024751178Z-f74fe604` | PASS — `max-actions-budget`, actions=40 | [summary](./phase4b-rereview5/20261112T024751178Z-f74fe604/summary.json) / [transitions](./phase4b-rereview5/20261112T024751178Z-f74fe604/transitions.jsonl) |
| Modal + HWND backtrack | `20261112T024837101Z-d4723e31` | PASS — `modalDialog`×2 + `closedWindow`×2 via HWND WindowPattern | [summary](./phase4b-rereview5/20261112T024837101Z-d4723e31/summary.json) / [transitions](./phase4b-rereview5/20261112T024837101Z-d4723e31/transitions.jsonl) |
| Non-modal second-window crawl | `20261112T024849579Z-1206ebc4` | PASS — `Open Second`=`newWindow`, then `Ping Second`=`success`, HWND close via WindowPattern | [summary](./phase4b-rereview5/20261112T024849579Z-1206ebc4/summary.json) / [transitions](./phase4b-rereview5/20261112T024849579Z-1206ebc4/transitions.jsonl) |

## Exit gate checklist / coverage matrix

| Check | Evidence |
|---|---|
| Complete fixture traversal | Dry-run candidates in [dry-run.json](./phase4b-rereview5/dry-run.json) |
| True modal | `modalDialog` + `New window(s): DA Explorer Dialog` in modal transitions |
| Non-modal second window crawl | `newWindow` for Open Second + Ping Second success on `DA Explorer Second` (not staleRef-only) |
| HWND-bound WindowPattern backtrack | `backtrack:close:0x…` / detail contains `WindowPattern`; close args are hwnd+processId (driver method `windowPattern`) |
| Post-switch HWND verify | Coordinator requires snapshot HWND == requested HWND before enqueue/`didSwitchToChild` |
| Window identity = HWND (+ PID validate) | Diff by HWND even when PID missing; missing-PID on new windows fails closed |
| Session pin fail-closed | Live session required; snapshot≠live rejected; pinned session enforced for root/child/new-window |
| NewWindow ≠ NoOp | Unit + live modal/second-window paths switch verified HWND then re-observe |
| Budget / cancel labeling | Empty-frontier path checks budget/cancel before `frontier-exhausted`; failed/timed-out closes count toward `maxActions` |
| External-window rejection | Detector + identity unit coverage; inaccessible/elevated (missing PID) fail closed |
| Schema enforcement | Runtime schema rejects misspelled budgets |

## Agent suite

`Phase4bSafeExplorerTests`: **35** passed locally before push.
