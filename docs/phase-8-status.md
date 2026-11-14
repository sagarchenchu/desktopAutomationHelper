# Phase 8 status

## Required final report

Phase 8 status: **WINDOWS ACCEPTANCE PASS** — DNU-23 S1–S9 re-run complete. Prior 9/9 PASS on `28b8712` remains retracted (HTTP inspect/stop not OK). PR #148 remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Review baseline: `212e270`
- Product review fixes: `8bf785f`
- CI assertion fix (green): `ab05a0f0a60b6f7f1473995ce75cb4eaec7f3acf` / `ab05a0f`
- Tip tested / evidence tip: `bac8a1f` (ab05a0f + highlight STA fix + acceptance evidence; see HEAD after tip-pointer commit)
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 — **draft** (do not merge)
- CI: **green** on `ab05a0f` (Build and Test + Publish Standalone Packages)

### Windows acceptance (DNU-23)

- Overall: **PASS** (9/9)
- Runtime base: `ab05a0f` + highlight STA overlay fix (committed with this evidence)
- HTTP gates asserted: `startHttpOk` / `inspectHttpOk` / `verifyHttpOk` / `stopHttpOk` where used
- Scenario 5: real CustomDrawnCanvas (empty Name/AutomationId) + `Vision:Enabled` gated `/debug/inspect` fallback (`AllowedUiaInaccessible`, vision candidates ≥ 1)

| # | Result | HTTP notes |
|---|---|---|
| 1 | PASS | start/inspect/verify/stop OK |
| 2 | PASS | start/inspect/verify/stop OK |
| 3 | PASS | start/inspect/verify/stop OK |
| 4 | PASS | start/inspect/stop OK (verify N/A); clickuia under yellow highlight |
| 5 | PASS | start/inspect/verify/stop OK + live inaccessible vision |
| 6 | PASS | `/vision/status` + `/vision/candidates` default-off |
| 7 | PASS | stale/DPI/layout/destructive + Delete label deny |
| 8 | PASS | sensitive screenshot not persisted |
| 9 | PASS | `/debug/hover` mouse probe |

Evidence: `docs/acceptance/phase8-*` (+ store mirror under Project `docs/acceptance/`).

### Exit gate

Windows re-acceptance **PASS**. PR stays **draft**; do not merge unless instructed.
