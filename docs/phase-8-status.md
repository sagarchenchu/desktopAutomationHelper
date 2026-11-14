# Phase 8 status

## Required final report

Phase 8 status: **CHANGES REQUESTED → fixes pushed** — P1/P2 from user review of draft PR #148 @ `212e270` addressed on branch. CI/unit tests re-run on this tip. Windows S1–S9 acceptance **PENDING** (DNU-23); prior 9/9 PASS on `28b8712` **retracted** (HTTP inspect/stop not OK). PR remains **draft**. Do **not** merge.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base: `origin/main`
- Review baseline: `212e270`
- Tip: see HEAD after review-fix commit
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 — **draft** (do not merge)

### Review fixes (P1/P2)

| # | Item | Status |
|---|---|---|
| P1-1 | Config-gated production vision (`Vision:*`, `/vision/*`, Debug inspect fallback); CustomDrawnCanvas empty Name; Hover/DoubleClick HTTP surfaces | Fixed |
| P1-2 | CoordinateActionGuard: click point + full candidate bounds inside window; HWND+PID fail-closed | Fixed |
| P1-3 | Destructive Candidate.Label heuristics denied even for Action=click | Fixed |
| P1-4 | Retract 9/9 PASS; acceptance asserts HTTP start/inspect/verify/stop OK (`phase8-assert-http-ok.ps1`) | Fixed (Windows re-run = DNU-23) |
| P2-5 | Live probe unavailable / zero-match never becomes unique via stale capture | Fixed |
| P2-6 | FlaUIUiaElementProbe captures real DpiScale via GetDpiForWindow | Fixed |

### Debug / vision surfaces

- Debug: `/debug/status|start|stop|toggle|inspect|verify|hover|doubleclick-allowed`
- Vision (default off): `/vision/status|evaluate|candidates|authorize|hover|doubleclick`
- `Vision:Enabled=false` in `appsettings.json`; bind via `IOptions<VisionFallbackOptions>`

### Windows acceptance

**PENDING (DNU-23).** Do not claim PASS until Windows worker re-runs S1–S9 on post-fix tip with `startHttpOk`/`inspectHttpOk`/`verifyHttpOk`/`stopHttpOk` all true. Helper: `docs/acceptance/phase8-assert-http-ok.ps1`.

### Known limitations

- Real OCR providers not enabled (fake only)
- No headless claim; interactive desktop required
- Vision independently gated; Debug Mode does not enable vision by itself

### Exit gate

Debug capture and visual fallback must work on controlled fixtures without intercepting clicks, persisting sensitive screenshots by default, or promoting any candidate automatically. **Code fixes in; Windows re-acceptance pending.** PR stays draft; do not merge unless instructed.
