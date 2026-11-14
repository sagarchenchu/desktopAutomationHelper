# Phase 8 status

## Required final report

Phase 8 status: **WINDOWS ACCEPTANCE PASS** - DNU-25 authorize-path S1-S9 on tipSha `dfd65d0`. Prior PASS claims retracted after round-3. PR #148 remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Round-3 baseline: `554daad`
- Product tip (tested / tipSha): `dfd65d040ccb5eda2a856f54b5a53793a3e6cdfa` / `dfd65d0`
- Evidence HEAD: see branch tip after DNU-25 evidence commit (tipSha in result files remains `dfd65d0`)
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 - **draft** (do not merge)

### Windows acceptance (DNU-25)

- Overall: **PASS**
- tipSha consistency: **PASS** - all `docs/acceptance/phase8-*-result.json` + summary + acceptance md record `dfd65d040ccb5eda2a856f54b5a53793a3e6cdfa` (no post-hoc rewrite)
- HTTP gates: start/inspect/verify/stop OK where used (S1-S5, S9)
- S5: candidates -> operator approve (`X-Operator-Token` from startup log) -> authorize Allowed=`True`
- S6: `/vision/authorize` denied when vision off
- S7: vision enabled; destructive reason: `Destructive or unknown actions cannot be approved.`

| # | Result | HTTP notes |
|---|---|---|
| 1 | PASS | start/inspect/verify/stop |
| 2 | PASS | start/inspect/verify/stop |
| 3 | PASS | start/inspect/verify/stop |
| 4 | PASS | start/inspect/stop (verify N/A) |
| 5 | PASS | candidates->approve->authorize Allowed=true |
| 6 | PASS | status/candidates off + authorize denied |
| 7 | PASS | guards + destructive after approve (vision on) |
| 8 | PASS | sensitive not persisted |
| 9 | PASS | /debug/hover |

### Exit gate

Windows re-acceptance PASS on single tipSha. PR stays draft; do not merge unless instructed.