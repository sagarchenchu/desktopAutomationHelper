# Phase 8 - Windows acceptance

- Tip tested (GitHub-resolvable): [`f0b222905055cad7d02e03fc0b280123ce79a531`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/f0b222905055cad7d02e03fc0b280123ce79a531) (`f0b2229`)
- tipSha (single, unchanged across all result files): `f0b222905055cad7d02e03fc0b280123ce79a531`
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft** - not merged)
- When: 2026-11-14T04:59:51.6344370+00:00
- Overall: **PASS**
- Note: DNU-27 after round-5 (`c62bec3`+). Real pixel capture; content-change-without-resize deny; S5 interactive presence (not ACL) -> authorize Allowed=true + consume fail; S7 destructive deny with vision enabled.

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
| 5 | server HWND pixel candidates -> interactive presence approve -> execution-token authorize -> **Allowed=true** |
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

- S5 pixelCapturePath: `ServerVisionCaptureService -> ScreenWindowPixelCapture PrintWindow client PNG (not HWND/PID text); restore foreground after presence before re-capture`
- S5 operatorPresenceSource: `Driver InteractiveOperatorPresenceService MessageBox on interactive desktop (title contains "Vision approval"); Yes clicked via UI Automation. ACL file self-read disabled / not used.`
- S5 contentChangeDenied: `True` reason: `Approval revalidation failed: content fingerprint changed.`
- S5 authorizeAllowed: `True` reason: `Authorized.`
- S5 tokenRebindOk (old token denied): `True` reason: `Unknown, expired, consumed, or mismatched execution token.`
- S5 consumeOk (second authorize denied): `True` reason: `Unknown, expired, consumed, or mismatched execution token.`
- S7 destructiveLabelReason: `Destructive or unknown actions cannot be approved.`