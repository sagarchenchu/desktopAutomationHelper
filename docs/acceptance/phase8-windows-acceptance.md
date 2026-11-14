# Phase 8 - Windows acceptance

- Tip tested (GitHub-resolvable): [`7d0bcd66c62826e201c591b8579d88f540029886`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/7d0bcd66c62826e201c591b8579d88f540029886) (`7d0bcd6`)
- tipSha (single, unchanged across all result files): `7d0bcd66c62826e201c591b8579d88f540029886`
- Branch: `feature/driver-phase-8-debug-vision`
- PR: #148 (**draft** - not merged)
- When: 2026-11-14T01:12:20.0733728+00:00
- Overall: **PASS**
- Note: DNU-24 SHA-consistent re-run after round-2 fixes (`7d0bcd6`+). HTTP start/inspect/verify/stop must be OK. S5 uses real CustomDrawnCanvas with `Vision:Enabled=true`. S6 asserts `/vision/authorize` denied when vision off.

## HTTP success requirement

| Flag | Endpoint |
|---|---|
| `startHttpOk` | `POST /debug/start` → 2xx |
| `inspectHttpOk` | `POST /debug/inspect` → 2xx |
| `verifyHttpOk` | `POST /debug/verify` → 2xx (when verify is part of the scenario) |
| `stopHttpOk` | `POST /debug/stop` → 2xx |

Helper: [`phase8-assert-http-ok.ps1`](phase8-assert-http-ok.ps1) (`-AssertVisionAuthorizeDenied` for S6).

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | Debug Mode start/stop + inspect on WinForms test app | Highlight + inspect without click; HTTP OK | PASS | `phase8-01-winforms-debug-result.json` |
| 2 | Debug Mode on WPF smoke app | Same as above; HTTP OK | PASS | `phase8-02-wpf-debug-result.json` |
| 3 | Candidate locator unique verify | Unique OK; no auto-promote; HTTP OK | PASS | `phase8-03-candidate-verify-result.json` |
| 4 | Yellow highlight does not intercept clicks | Click reaches underlying control; HTTP OK | PASS | `phase8-04-highlight-passthrough-result.json` |
| 5 | CustomDrawnCanvas vision path (Enabled=true) | Real inaccessible canvas → vision candidates; no auto-promote | PASS | `phase8-05-customdrawn-vision-result.json` |
| 6 | Vision disabled by default | No vision candidates; `/vision/authorize` denied | PASS | `phase8-06-vision-default-off-result.json` |
| 7 | Stale/DPI/layout change | Coordinate action denied | PASS | `phase8-07-stale-dpi-guards-result.json` |
| 8 | Sensitive screenshot | Not persisted by default | PASS | `phase8-08-sensitive-screenshot-result.json` |
| 9 | Hover probe | Mouse probe via `/debug/hover`; not UIA capability | PASS | `phase8-09-hover-probe-result.json` |