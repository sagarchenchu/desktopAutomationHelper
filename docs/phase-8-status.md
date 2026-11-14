# Phase 8 status

## Required final report

Phase 8 status: **WINDOWS ACCEPTANCE PASS** - DNU-26 token-safety S1-S9 on tipSha `f9913ef`. Prior DNU-25 PASS superseded. PR #148 remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Round-4 baseline (review): `7448395`
- Round-4 product tip: `a647efd306cbdded76c7cb8bc45e7bd80fbc24a5` / `a647efd`
- CI green on round-4 docs tip: `2531167` (4/4) before ACL delivery
- Product tip (tested / tipSha): `f9913efc1010e3ca3f4641ada4ce2d84c445433b` / `f9913ef` (ACL operator secret + round-4)
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 - **draft** (do not merge)

### Round-4 fixes (+ ACL delivery on tipSha)

- Immutable one-use execution token grants; re-approval drops prior grant (no orphans); atomic claim before Allowed=true; submitted token verified against issued
- Approve re-captures window content and revalidates fingerprint/label/bounds before issuing token
- Server capture binds cursor UIA tree HWND/PID to requested window; cursor must be inside live bounds
- Operator token never written to startup banner/logs; delivered via current-user ACL secret file

### Windows acceptance (DNU-26)

- Overall: **PASS**
- tipSha consistency: **PASS** - all `docs/acceptance/phase8-*-result.json` + summary + acceptance md record `f9913efc1010e3ca3f4641ada4ce2d84c445433b` (no post-hoc rewrite)
- HTTP gates: start/inspect/verify/stop OK where used (S1-S5, S9)
- Operator token source: ACL-protected secret (never startup log)
- S5: authorize Allowed=`True`; rebind old-token denied=`True`; second authorize denied=`True`
- S6: `/vision/authorize` denied when vision off
- S7: vision enabled; destructive reason: `Destructive or unknown actions cannot be approved.`

| # | Result | HTTP notes |
|---|---|---|
| 1 | PASS | start/inspect/verify/stop |
| 2 | PASS | start/inspect/verify/stop |
| 3 | PASS | start/inspect/verify/stop |
| 4 | PASS | start/inspect/stop (verify N/A) |
| 5 | PASS | ACL approve/rebind/authorize Allowed=true + consume fail |
| 6 | PASS | status/candidates off + authorize denied |
| 7 | PASS | guards + destructive after approve (vision on) |
| 8 | PASS | sensitive not persisted |
| 9 | PASS | /debug/hover |

### Exit gate

Windows DNU-26 PASS on single tipSha. PR stays draft; do not merge unless instructed.
