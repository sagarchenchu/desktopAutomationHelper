# Phase 8 - Windows acceptance

- Tip tested: **PENDING** — DNU-27 after round-5 (real pixels, original baseline approve, interactive presence, PID fail-closed)
- tipSha: *(fill on DNU-27 run — must match product tip)*
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft** - not merged)
- Overall: **PENDING**
- Note: Prior DNU-26 PASS on `f9913ef` superseded. Approve requires driver interactive presence (ACL file self-read is not human approval).

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
| 1 | Debug Mode start/stop + inspect on WinForms test app | Highlight + inspect without click; HTTP OK | PENDING | DNU-27 |
| 2 | Debug Mode on WPF smoke app | Same as above; HTTP OK | PENDING | DNU-27 |
| 3 | Candidate locator unique verify | Unique OK; no auto-promote; HTTP OK | PENDING | DNU-27 |
| 4 | Yellow highlight does not intercept clicks | Click reaches underlying control; HTTP OK | PENDING | DNU-27 |
| 5 | CustomDrawnCanvas authorize path (Enabled=true) | candidates->presence approve->authorize Allowed=true | PENDING | DNU-27 |
| 6 | Vision disabled by default | No vision candidates; `/vision/authorize` denied | PENDING | DNU-27 |
| 7 | Stale/DPI + destructive after approve | deny with vision on | PENDING | DNU-27 |
| 8 | Sensitive screenshot | Not persisted by default | PENDING | DNU-27 |
| 9 | Hover probe | Mouse probe via `/debug/hover`; not UIA capability | PENDING | DNU-27 |

## S5 / S7 evidence highlights

- Prior DNU-26 evidence superseded; do not treat JSON result files as current tip PASS.
- S5 approval: driver interactive presence (not ACL file / not automation self-read).
- S7: destructive deny after approve with vision enabled.
