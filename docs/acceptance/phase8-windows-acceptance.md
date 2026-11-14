# Phase 8 - Windows acceptance

- Tip tested (GitHub-resolvable): [`f9913efc1010e3ca3f4641ada4ce2d84c445433b`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/f9913efc1010e3ca3f4641ada4ce2d84c445433b) (`f9913ef`)
- tipSha (single, unchanged across all result files): `f9913efc1010e3ca3f4641ada4ce2d84c445433b`
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft** - not merged)
- When: 2026-11-14T02:30:22.0582373+00:00
- Overall: **PASS**
- Note: DNU-26 token-safety re-run after round-4 (`a647efd`+). S5 ACL operator secret -> approve/rebind/authorize Allowed=true + consume fail. S7 destructive deny with vision enabled. Operator token from ACL file (never startup log).

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
- S5 operatorTokenSource: `ACL-protected secret file (C:\Users\ADMIN\AppData\Local\DesktopAutomationDriver\operator-token); never startup log / not automation bearer / not GET /verify`
- S5 tokenRebindOk (old token denied): `True` reason: `Unknown, expired, consumed, or mismatched execution token.`
- S5 consumeOk (second authorize denied): `True` reason: `Unknown, expired, consumed, or mismatched execution token.`
- S7 destructiveLabelReason: `Destructive or unknown actions cannot be approved.`