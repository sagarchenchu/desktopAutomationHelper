# Phase 8 - Windows acceptance

- Prior tips (`28b8712`, `ab05a0f`/`464ab9e` DNU-23 evidence) are **not** authoritative for the current review tip — SHA inconsistency retracted.
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft / CHANGES REQUESTED → round-2 fixes** — not merged)
- Overall: **PENDING** (DNU-24 Windows re-run). Do **not** claim 9/9 PASS until DNU-24 lands on one clean tip SHA with every result file recording that unchanged SHA.

## HTTP success requirement

| Flag | Endpoint |
|---|---|
| `startHttpOk` | `POST /debug/start` → 2xx |
| `inspectHttpOk` | `POST /debug/inspect` → 2xx |
| `verifyHttpOk` | `POST /debug/verify` → 2xx (when verify is part of the scenario) |
| `stopHttpOk` | `POST /debug/stop` → 2xx |

Helper: [`phase8-assert-http-ok.ps1`](phase8-assert-http-ok.ps1).

Vision Scenario 6 must also assert `POST /vision/authorize` is **denied** when `Vision:Enabled=false` (`authorizeHttpDenied=true`).

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | Debug Mode start/stop + inspect on WinForms test app | Highlight + inspect without click; HTTP OK | PENDING | `phase8-01-winforms-debug-result.json` |
| 2 | Debug Mode on WPF smoke app | Same as above; HTTP OK | PENDING | `phase8-02-wpf-debug-result.json` |
| 3 | Candidate locator unique verify | Unique OK; no auto-promote; HTTP OK | PENDING | `phase8-03-candidate-verify-result.json` |
| 4 | Yellow highlight does not intercept clicks | Click reaches underlying control; HTTP OK | PENDING | `phase8-04-highlight-passthrough-result.json` |
| 5 | CustomDrawnCanvas vision path (Enabled=true) | Real inaccessible canvas → server-issued vision candidates; no auto-promote | PENDING | `phase8-05-customdrawn-vision-result.json` |
| 6 | Vision disabled by default | No candidates; **authorize denied** when off | PENDING | `phase8-06-vision-default-off-result.json` |
| 7 | Stale/DPI/layout + destructive/unknown | Coordinate action denied | PENDING | `phase8-07-stale-dpi-guards-result.json` |
| 8 | Sensitive screenshot | Not persisted by default | PENDING | `phase8-08-sensitive-screenshot-result.json` |
| 9 | Hover probe | Mouse probe via `/debug/hover`; not UIA capability | PENDING | `phase8-09-hover-probe-result.json` |
