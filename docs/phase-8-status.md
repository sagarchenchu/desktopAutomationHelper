# Phase 8 status

## Required final report

Phase 8 status: **WINDOWS ACCEPTANCE PASS** - DNU-29 confirmation-honesty S1-S9 on tipSha `8eeef92`. Prior PASS claims superseded. PR #148 remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Round-7 product tip: `7c22673d497e22a9ba6e6a1a710f4f79f8ef7951` / `7c22673`
- Tested tipSha (contains product): `8eeef9290e45342f5981cbb492dffd4eda840eed` / `8eeef92`
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 - **draft** (do not merge)

### Windows acceptance (DNU-29)

- Overall: **PASS**
- tipSha consistency: **PASS** - all `docs/acceptance/phase8-*-result.json` + summary + acceptance md record `8eeef9290e45342f5981cbb492dffd4eda840eed` (no post-hoc rewrite)
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

Windows DNU-29 PASS on single tipSha. PR stays draft; do not merge unless instructed.