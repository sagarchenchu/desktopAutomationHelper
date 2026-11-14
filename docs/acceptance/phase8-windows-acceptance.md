# Phase 8 - Windows acceptance

- Prior PASS claims (including DNU-23 / DNU-24 evidence) are **retracted** as incomplete for the current security model (operator channel, server capture, execution tokens).
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft / CHANGES REQUESTED → round-3 fixes** — not merged)
- Overall: **PENDING** (DNU-25 Windows re-run). Do **not** claim 9/9 PASS until DNU-25 lands on one clean tip SHA.

## HTTP / vision success requirements

Debug HTTP: `startHttpOk` / `inspectHttpOk` / `verifyHttpOk` / `stopHttpOk` via [`phase8-assert-http-ok.ps1`](phase8-assert-http-ok.ps1).

Vision (DNU-25 must exercise):

| Scenario | Requirement |
|---|---|
| 5 | `candidates` (server HWND capture) → `approve` (`X-Operator-Token` + action) → `authorize` (execution token) → **Allowed=true** |
| 6 | Vision off: candidates empty; authorize denied |
| 7 | Vision **enabled**: destructive label rejection **after** operator approve attempt (not “Vision fallback disabled”) |

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | WinForms debug | HTTP OK | PENDING | `phase8-01-winforms-debug-result.json` |
| 2 | WPF debug | HTTP OK | PENDING | `phase8-02-wpf-debug-result.json` |
| 3 | Candidate verify | HTTP OK | PENDING | `phase8-03-candidate-verify-result.json` |
| 4 | Highlight passthrough | HTTP OK | PENDING | `phase8-04-highlight-passthrough-result.json` |
| 5 | CustomDrawn vision success path | candidates→approve→authorize Allowed=true | PENDING | `phase8-05-customdrawn-vision-result.json` |
| 6 | Vision default off | authorize denied when off | PENDING | `phase8-06-vision-default-off-result.json` |
| 7 | Stale/DPI + destructive after approve | deny with vision on | PENDING | `phase8-07-stale-dpi-guards-result.json` |
| 8 | Sensitive screenshot | not persisted | PENDING | `phase8-08-sensitive-screenshot-result.json` |
| 9 | Hover probe | `/debug/hover` | PENDING | `phase8-09-hover-probe-result.json` |
