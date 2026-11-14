# Phase 8 status

## Required final report

Phase 8 status: **CHANGES REQUESTED → round-2 fixes pushed** — P1/P2 from user review of draft PR #148 @ `785304f` addressed on branch. Prior Windows 9/9 PASS claims **retracted** (SHA inconsistency + security gaps). Windows S1–S9 re-run is **DNU-24** (queued). PR remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Round-2 review baseline: `785304f`
- Tip: `7d0bcd66c62826e201c591b8579d88f540029886` / `7d0bcd6`
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 — **draft** (do not merge)

### Round-2 review fixes

| # | Item | Status |
|---|---|---|
| P1-1 | Gate `AuthorizeCoordinateAction` / `POST /vision/authorize` on `Vision:Enabled`; S6 asserts authorize deny when off | Fixed |
| P1-2 | Opaque server-issued candidate id; server clock; stored approval; fresh Win32/UIA HWND/PID/bounds/DPI probe; reject unapproved/auto-promoted; no trust of client window/timestamps | Fixed |
| P1-3 | Retract inconsistent acceptance SHAs; leave Windows re-run to DNU-24; do not claim 9/9 PASS | Fixed |
| P2 | Fail-closed destructive heuristics: Send/Approve/Publish/Install labels + unknown action strings denied | Fixed |

### Vision auth surfaces

- Issue: `POST /vision/candidates` → opaque `vc-…` ids (`requiresHumanApproval=true`)
- Approve: `POST /vision/approve` `{ candidateId }`
- Authorize: `POST /vision/authorize` `{ candidateId, action }` — gated on `Vision:Enabled`, approval, fresh probe

### Windows acceptance

**PENDING (DNU-24).** Do not claim PASS. One clean tip SHA; every result file must record that SHA. S6 must include authorize deny when vision off (`-AssertVisionAuthorizeDenied`).

### Exit gate

Code fixes in; Windows re-acceptance pending DNU-24. PR stays draft; do not merge unless instructed.
