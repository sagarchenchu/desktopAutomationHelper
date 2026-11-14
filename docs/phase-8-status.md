# Phase 8 status

## Required final report

Phase 8 status: **WINDOWS ACCEPTANCE PENDING** — DNU-26 after round-4 security fixes. Prior DNU-25 PASS on `dfd65d0` is **superseded / not claimed** for current tip. PR #148 remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Round-4 baseline (review): `7448395`
- Round-4 product tip: `a647efd306cbdded76c7cb8bc45e7bd80fbc24a5` / `a647efd`
- Branch HEAD / CI tip: `25311678a72bb715cf59f3d8ec3929c751910c62` / `2531167`
- CI: **green** (4/4) on `2531167`
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 — **draft** (do not merge)

### Round-4 fixes (this tip)

- Immutable one-use execution token grants; re-approval drops prior grant (no orphans); atomic claim before Allowed=true; submitted token verified against issued
- Approve re-captures window content and revalidates fingerprint/label/bounds before issuing token
- Server capture binds cursor UIA tree HWND/PID to requested window; cursor must be inside live bounds
- Operator token never written to startup banner/logs (`StartupBanner`)

### Windows acceptance (DNU-26)

- Overall: **PENDING**
- Ready for Windows: **yes** (CI green on tip `2531167`)
- tipSha for DNU-26: `25311678a72bb715cf59f3d8ec3929c751910c62`
- Operator token via secure UI/ACL (not startup log)

### Exit gate

Keep draft. Windows DNU-26 required before any PASS claim. Do not merge unless instructed.
