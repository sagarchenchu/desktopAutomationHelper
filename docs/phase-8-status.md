# Phase 8 status

## Required final report

Phase 8 status: **WINDOWS ACCEPTANCE PASS** - DNU-27 human-approve / pixel re-capture S1-S9 on tipSha `f0b2229`. Prior DNU-26 PASS superseded. PR #148 remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Round-5 product baseline: `c62bec3`
- Product tip (tested / tipSha): `f0b222905055cad7d02e03fc0b280123ce79a531` / `f0b2229`
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 - **draft** (do not merge)

### Windows acceptance (DNU-27)

- Overall: **PASS**
- tipSha consistency: **PASS** - all `docs/acceptance/phase8-*-result.json` + summary + acceptance md record `f0b222905055cad7d02e03fc0b280123ce79a531` (no post-hoc rewrite)
- HTTP gates: start/inspect/verify/stop OK where used (S1-S5, S9)
- Pixel path: ScreenWindowPixelCapture / PrintWindow client PNG (not HWND/PID text)
- Approval: driver interactive presence MessageBox (ACL file self-read disabled)
- S5: authorize Allowed=`True`; content-change deny=`True`; rebind old-token denied=`True`; second authorize denied=`True`
- S6: `/vision/authorize` denied when vision off
- S7: vision enabled; destructive reason: `Destructive or unknown actions cannot be approved.`

| # | Result | HTTP notes |
|---|---|---|
| 1 | PASS | start/inspect/verify/stop |
| 2 | PASS | start/inspect/verify/stop |
| 3 | PASS | start/inspect/verify/stop |
| 4 | PASS | start/inspect/stop (verify N/A) |
| 5 | PASS | presence approve/rebind/authorize Allowed=true + consume fail |
| 6 | PASS | status/candidates off + authorize denied |
| 7 | PASS | guards + destructive after approve (vision on) |
| 8 | PASS | sensitive not persisted |
| 9 | PASS | /debug/hover |

### Exit gate

Windows DNU-27 PASS on single tipSha. PR stays draft; do not merge unless instructed.