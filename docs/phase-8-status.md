# Phase 8 status

## Required final report

Phase 8 status: **WINDOWS ACCEPTANCE PASS** - DNU-24 SHA-consistent S1-S9 on tipSha `7d0bcd6`. Prior DNU-23 PASS retracted after round-2. PR #148 remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Round-2 baseline: `785304f`
- Product tip (tested / tipSha): `7d0bcd66c62826e201c591b8579d88f540029886` / `7d0bcd6`
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 - **draft** (do not merge)

### Windows acceptance (DNU-24)

- Overall: **PASS** (9/9)
- tipSha consistency: **PASS** - all `docs/acceptance/phase8-*-result.json` + `phase8-windows-summary.json` + acceptance md record `7d0bcd66c62826e201c591b8579d88f540029886` (no post-hoc rewrite)
- HTTP gates: start/inspect/verify/stop OK where used (S1-S5, S9)
- S5: real CustomDrawnCanvas + `Vision:Enabled` gated opaque candidate path (`AllowedUiaInaccessible`, candidates >= 1)
- S6: `/vision/status` off, `/vision/candidates` empty/`DisabledByDefault`, `/vision/authorize` denied (`Vision:Enabled=false`)

| # | Result | HTTP notes |
|---|---|---|
| 1 | PASS | start/inspect/verify/stop OK |
| 2 | PASS | start/inspect/verify/stop OK |
| 3 | PASS | start/inspect/verify/stop OK |
| 4 | PASS | start/inspect/stop OK (verify N/A) |
| 5 | PASS | start/inspect/verify/stop OK + live inaccessible vision |
| 6 | PASS | status/candidates off + authorize denied |
| 7 | PASS | stale/DPI/layout/destructive + opaque-id authorize deny |
| 8 | PASS | sensitive screenshot not persisted |
| 9 | PASS | `/debug/hover` mouse probe |

Evidence: `docs/acceptance/phase8-*` (+ store mirror under Project `docs/acceptance/`).

### Exit gate

Windows re-acceptance **PASS** on single tipSha `7d0bcd6`. PR stays **draft**; do not merge unless instructed.
