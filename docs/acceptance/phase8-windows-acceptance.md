# Phase 8 - Windows acceptance

- Tip previously exercised: [`28b8712db032e867bb8433f1f2bd69161d85004d`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/28b8712db032e867bb8433f1f2bd69161d85004d) (`28b8712`) — **HTTP inspect/verify/stop were NOT OK**; overall PASS claim **retracted**.
- Product tip after review fixes: see PR #148 HEAD (post P1/P2). Windows S1–S9 re-run is **DNU-23** (follow-on); do **not** claim 9/9 PASS until that re-run lands with `startHttpOk`/`inspectHttpOk`/`verifyHttpOk`/`stopHttpOk` all true.
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft / CHANGES REQUESTED → fixes pushed** — not merged)
- Overall: **PENDING** Windows re-acceptance (DNU-23)

## HTTP success requirement

Every Debug Mode scenario that exercises the driver HTTP surface **must** assert:

| Flag | Endpoint |
|---|---|
| `startHttpOk` | `POST /debug/start` → 2xx |
| `inspectHttpOk` | `POST /debug/inspect` → 2xx |
| `verifyHttpOk` | `POST /debug/verify` → 2xx (when verify is part of the scenario) |
| `stopHttpOk` | `POST /debug/stop` → 2xx |

Helper: [`phase8-assert-http-ok.ps1`](phase8-assert-http-ok.ps1). On-disk `candidate.json` alone is **not** sufficient for PASS.

Vision scenarios additionally use config-gated `GET/POST /vision/*` (`Vision:Enabled=true` only for Scenario 5).

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | Debug Mode start/stop + inspect on WinForms test app | Highlight + inspect without click; HTTP OK | PENDING | `phase8-01-winforms-debug-result.json` (stale; re-run required) |
| 2 | Debug Mode on WPF smoke app | Same as above; HTTP OK | PENDING | `phase8-02-wpf-debug-result.json` |
| 3 | Candidate locator unique verify | Unique OK; no auto-promote; HTTP OK | PENDING | `phase8-03-candidate-verify-result.json` |
| 4 | Yellow highlight does not intercept clicks | Click reaches underlying control | PENDING | `phase8-04-highlight-passthrough-result.json` |
| 5 | CustomDrawnCanvas vision path (Enabled=true) | Real app inaccessible canvas → vision candidates; no auto-promote | PENDING | `phase8-05-customdrawn-vision-result.json` |
| 6 | Vision disabled by default | No vision candidates | PENDING | `phase8-06-vision-default-off-result.json` |
| 7 | Stale/DPI/layout change | Coordinate action denied | PENDING | `phase8-07-stale-dpi-guards-result.json` |
| 8 | Sensitive screenshot | Not persisted by default | PENDING | `phase8-08-sensitive-screenshot-result.json` |
| 9 | Hover probe | Mouse probe via `/debug/hover`; not UIA capability | PENDING | `phase8-09-hover-probe-result.json` |
