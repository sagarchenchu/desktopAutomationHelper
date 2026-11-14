# Phase 8 - Windows acceptance

- Tip tested: **PENDING** — DNU-26 after round-4 (atomic tokens, approve re-capture/fingerprint, cursor HWND bind, no operator token in logs)
- tipSha: *(fill on DNU-26 run — must match product tip)*
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft** - not merged)
- Overall: **PENDING**
- Note: Prior DNU-25 PASS on `dfd65d0` superseded. Operator token must **not** come from startup logs (use interactive/ACL secret).

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
| 1 | Debug Mode start/stop + inspect on WinForms test app | Highlight + inspect without click; HTTP OK | PENDING | DNU-26 |
| 2 | Debug Mode on WPF smoke app | Same as above; HTTP OK | PENDING | DNU-26 |
| 3 | Candidate locator unique verify | Unique OK; no auto-promote; HTTP OK | PENDING | DNU-26 |
| 4 | Yellow highlight does not intercept clicks | Click reaches underlying control; HTTP OK | PENDING | DNU-26 |
| 5 | CustomDrawnCanvas authorize path (Enabled=true) | candidates->approve->authorize Allowed=true | PENDING | DNU-26 |
| 6 | Vision disabled by default | No vision candidates; `/vision/authorize` denied | PENDING | DNU-26 |
| 7 | Stale/DPI + destructive after approve | deny with vision on | PENDING | DNU-26 |
| 8 | Sensitive screenshot | Not persisted by default | PENDING | DNU-26 |
| 9 | Hover probe | Mouse probe via `/debug/hover`; not UIA capability | PENDING | DNU-26 |

## S5 / S7 evidence highlights

- Prior DNU-25 evidence on `dfd65d0` superseded; do not treat JSON result files as current tip PASS.
- S5 operatorTokenSource for DNU-26: interactive UI / ACL-protected secret (**never** startup log).
- S7: destructive deny after approve with vision enabled.