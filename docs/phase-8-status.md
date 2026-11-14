# Phase 8 status

## Required final report

Phase 8 status: **WINDOWS ACCEPTANCE PASS** - DNU-32 round-8 cancel-cleanup S1-S9 on tipSha `9c927ca`. Prior PASS claims superseded. PR #148 remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Round-8 product tip: `9c927cad21168afc584138a77b686c5362731640` / `9c927ca`
- Tested tipSha (contains product): `9c927cad21168afc584138a77b686c5362731640` / `9c927ca`
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 - **draft** (do not merge)

### Windows acceptance (DNU-32)

- Overall: **PASS**
- tipSha consistency: **PASS** - all `docs/acceptance/phase8-*-result.json` + summary + acceptance md record `9c927cad21168afc584138a77b686c5362731640` (no post-hoc rewrite)
- HTTP gates: start/inspect/verify/stop OK where used (S1-S5, S9)
- Pixel path: ScreenWindowPixelCapture PrintWindow client PNG (not HWND/PID text)
- API honesty: `requiresHumanApproval=false`, `requiresOperatorConfirmationUx=true` (apiHonestyOk=`True`)
- Approval: **operator confirmation UX** (UIA-automatable Yes); no human-approval / human-presence claim; ACL unused
- Timeout deny: `True` reason: `Operator confirmation timed out.`
- S5: authorize Allowed=`True`; content-change deny=`True`; rebind old-token denied=`True`; second authorize denied=`True`
- S6: `/vision/authorize` denied when vision off
- S7: vision enabled; destructive reason: `Destructive or unknown actions cannot be approved.`

| # | Result | HTTP notes |
|---|---|---|
| 1 | PASS | start/inspect/verify/stop |
| 2 | PASS | start/inspect/verify/stop |
| 3 | PASS | start/inspect/verify/stop + API honesty flags on artifact |
| 4 | PASS | start/inspect/stop (verify N/A) |
| 5 | PASS | API honesty + confirmation UX + timeout deny + authorize Allowed=true + consume |
| 6 | PASS | status/candidates off + authorize denied |
| 7 | PASS | guards + destructive after approve (vision on) |
| 8 | PASS | sensitive not persisted |
| 9 | PASS | /debug/hover |

### Exit gate

Windows DNU-32 PASS on single tipSha. PR stays draft; do not merge unless instructed.