# Phase 8 - Windows acceptance

- Tip tested (GitHub-resolvable): [`dfd65d040ccb5eda2a856f54b5a53793a3e6cdfa`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/dfd65d040ccb5eda2a856f54b5a53793a3e6cdfa) (`dfd65d0`)
- tipSha (single, unchanged across all result files): `dfd65d040ccb5eda2a856f54b5a53793a3e6cdfa`
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft** - not merged)
- When: 2026-11-14T02:08:00.3538550+00:00
- Overall: **PASS**
- Note: DNU-25 authorize-path re-run after round-3 (`dfd65d0`+). S5 candidates->operator approve->authorize Allowed=true. S7 destructive deny with vision enabled. Operator token from driver startup log (not automation bearer).

## HTTP / vision success requirements

| Flag | Endpoint |
|---|---|
| `startHttpOk` | `POST /debug/start` -> 2xx |
| `inspectHttpOk` | `POST /debug/inspect` -> 2xx |
| `verifyHttpOk` | `POST /debug/verify` -> 2xx (when verify is part of the scenario) |
| `stopHttpOk` | `POST /debug/stop` -> 2xx |

Helper: [`phase8-assert-http-ok.ps1`](phase8-assert-http-ok.ps1) (`-AssertVisionAuthorizeDenied` for S6).

| Scenario | Requirement |
|---|---|
| 5 | server HWND candidates -> `X-Operator-Token` approve -> execution-token authorize -> **Allowed=true** |
| 6 | Vision off: candidates empty; authorize denied |
| 7 | Vision **enabled**: destructive label rejection after operator approve attempt (not Vision:Enabled=false) |

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | Debug Mode start/stop + inspect on WinForms test app | Highlight + inspect without click; HTTP OK | PASS | `phase8-01-winforms-debug-result.json` |
| 2 | Debug Mode on WPF smoke app | Same as above; HTTP OK | PASS | `phase8-02-wpf-debug-result.json` |
| 3 | Candidate locator unique verify | Unique OK; no auto-promote; HTTP OK | PASS | `phase8-03-candidate-verify-result.json` |
| 4 | Yellow highlight does not intercept clicks | Click reaches underlying control; HTTP OK | PASS | `phase8-04-highlight-passthrough-result.json` |
| 5 | CustomDrawnCanvas authorize path (Enabled=true) | candidates->approve->authorize Allowed=true | PASS | `phase8-05-customdrawn-vision-result.json` |
| 6 | Vision disabled by default | No vision candidates; `/vision/authorize` denied | PASS | `phase8-06-vision-default-off-result.json` |
| 7 | Stale/DPI + destructive after approve | deny with vision on | PASS | `phase8-07-stale-dpi-guards-result.json` |
| 8 | Sensitive screenshot | Not persisted by default | PASS | `phase8-08-sensitive-screenshot-result.json` |
| 9 | Hover probe | Mouse probe via `/debug/hover`; not UIA capability | PASS | `phase8-09-hover-probe-result.json` |

## S5 / S7 evidence highlights

- S5 authorizeAllowed: `True` reason: `Authorized.`
- S5 operatorTokenSource: `driver startup log (X-Operator-Token / Operator token banner); not automation bearer / not GET /verify`
- S7 destructiveLabelReason: `Destructive or unknown actions cannot be approved.`