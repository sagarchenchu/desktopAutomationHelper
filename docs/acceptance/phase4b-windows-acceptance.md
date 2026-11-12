# Phase 4B Windows acceptance

Recorded on the self-hosted Windows worker against tip of
`feature/agent-phase-4b-safe-explorer` after the **sixth** re-review fixes.

- Base SHA: `2c6336042127a9553bf974f368f9dbd5af658fe1` (`origin/main`, Phase 4A merge #140)
- Fixture: `tools/DesktopAutomationExplorerFixture` (**DA Explorer Fixture**)
- **PR-auditable evidence:** [`phase4b-rereview6/`](./phase4b-rereview6/) (summaries, transitions, policies committed in-repo)

## Exit-gate runs (this tip)

| Run | explorationId | Result | Evidence |
|---|---|---|---|
| Dry-run | `20261112T060542511Z-5b049e2f` | PASS — dry-run completed | [summary](./phase4b-rereview6/20261112T060542511Z-5b049e2f/summary.json) / [stdout](./phase4b-rereview6/dry-run.json) |
| Bounded navigate | `20261112T060544139Z-9d62233f` | PASS — `max-actions-budget`, actions=40 | [summary](./phase4b-rereview6/20261112T060544139Z-9d62233f/summary.json) / [transitions](./phase4b-rereview6/20261112T060544139Z-9d62233f/transitions.jsonl) |
| Modal + HWND backtrack | `20261112T060740200Z-f6a6c903` | PASS — `modalDialog` + `closedWindow` via HWND WindowPattern | [summary](./phase4b-rereview6/20261112T060740200Z-f6a6c903/summary.json) / [transitions](./phase4b-rereview6/20261112T060740200Z-f6a6c903/transitions.jsonl) |
| Non-modal second-window crawl | `20261112T060745583Z-8ad995de` | PASS — `Open Second`=`newWindow`, `Ping Second`=`success`, HWND close via WindowPattern | [summary](./phase4b-rereview6/20261112T060745583Z-8ad995de/summary.json) / [transitions](./phase4b-rereview6/20261112T060745583Z-8ad995de/transitions.jsonl) |

## Exit gate checklist / coverage matrix

| Check | Evidence |
|---|---|
| Complete fixture traversal | Dry-run candidates in [dry-run.json](./phase4b-rereview6/dry-run.json) |
| True modal | `modalDialog` + `New window(s): DA Explorer Dialog` in modal transitions |
| Non-modal second window crawl | `newWindow` for Open Second + Ping Second success on `DA Explorer Second` |
| HWND-bound WindowPattern backtrack | `backtrack:close:0x…` / detail contains `WindowPattern` |
| HWND+PID switch | Coordinator passes `processId` with `switchwindow`; driver rejects PID mismatch before activation |
| HWND reuse detection | Matching HWND with different PID surfaces as remove+add (unit) |
| Window-aware frontier | Candidates carry `TargetHwnd`/`TargetProcessId`; retarget before exec (dual-sibling unit) |
| Budget during bind | Reports `max-actions-budget` / cancel, not `new-window-bind-failed` (unit) |
| Missing-PID fail-closed | Detector unit coverage |
| Session pin fail-closed | Live session required; pinned session enforced |

## Agent suite

`Phase4bSafeExplorerTests`: **38** passed locally before push.
