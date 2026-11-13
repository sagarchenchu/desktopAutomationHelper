# Phase 8 status

## Required final report

Phase 8 status: **READY FOR REVIEW** — CI **green** on tip `65a3326` (4/4); Windows acceptance **9/9 PASS**; `/debug/inspect` re-smoke **OK**; evidence on #148; #149 closed. PR remains **draft**. Do **not** merge unless explicitly instructed.

Branch/base/head/PR:
- Feature branch: `feature/driver-phase-8-debug-vision`
- Base SHA: `8a4a165394623400af7a2686828ce8a6b49be044` / `8a4a165` (`origin/main` after Phase 7 merge)
- Head / review tip: `65a332695d24489e492efd8417e0f3abd5445871` / `65a3326`
- Windows S1–S9 acceptance tip: `28b8712db032e867bb8433f1f2bd69161d85004d` / `28b8712`
- Inspect re-smoke: **OK** on tip `65a3326` (IntPtr hex `windowHandle`)
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/148 — **draft / READY FOR REVIEW** (do not merge unless instructed)

### CI failures / fixes (historical)

1. **`75649ee` → `5972d28`:** `FlaUIUiaElementProbe.cs` CS0121 — `Math.Round` cast to `(double)`.
2. **`5972d28` → `28b8712`:** `Phase8DebugVisionTests.cs` `using System.Drawing;` (CS0246/CS0535).
3. **`28b8712` → `af6756f`:** IntPtr STJ 500 — first `[JsonIgnore]` attempt; unit test round-trip failed (`Actual: "0"`).
4. **`af6756f` → `65a3326`:** `IntPtrHexJsonConverter` on `WindowHandle` (hex string in/out). CI green.

### Debug services

| Service | Location |
|---|---|
| `DebugSessionCoordinator` | `src/DesktopAutomationDriver/Services/Debug/` |
| `UiaTreeCaptureService` | same |
| `LocatorCandidateGenerator` | same |
| `NonInteractiveHighlightService` | same |
| `CandidateVerificationService` | same |
| `DebugArtifactWriter` | same |

Vision: `VisionFallbackPolicy`, `VisionScreenshotService`, `FakeVisionProvider`, `CoordinateActionGuard`, `VisionFallbackCoordinator` under `Services/Vision/`.

### Large class line changes

| File | Diff vs base |
|---|---|
| `RecordingOverlayWindow.cs` | **+37 / −2** (thin Ctrl+D + status/inspect forward only) |
| `UiService.cs` | **0** (unchanged) |

### Highlight behavior

Yellow click-through overlay (`WS_EX_TRANSPARENT` / no activate). `InterceptsClicks=false`. Cleared on Debug Mode stop / dispose.

### Candidate behavior

Phase-3 / object-repository locator strategies; unique verification required; artifacts under `debug-candidates/<session>/candidate.json` with `autoPromoted=false` and `requiresHumanApproval=true`. Never writes approved `pages/`.

### Vision enablement policy

Disabled by default (`Vision:Enabled=false` / `VisionFallbackOptions.Enabled=false`). Allowed only when enabled **and** (UIA inaccessible **or** `AllowWhenPolicyPermits`). Never replaces working UIA locators. Provider=`fake` in Phase 8.

### Stale/DPI safeguards

`CoordinateActionGuard` rejects stale screenshots (default max age 1500ms), window identity mismatch, DPI/scale change, bounds/layout change, and out-of-window bounds. Saved coordinates are not reused after layout change.

### Screenshot privacy

Window-scoped only (unrestricted desktop throws). Sensitive markers → not persisted by default (`PersistSensitiveScreenshots=false`). Optional text redaction helper.

### Agent tests

`Phase8VisionPolicyTests` — default-off, gates, provider validation, hover probe rules.

### Driver tests

`Phase8DebugVisionTests` — service separation, candidate quality, unique/ambiguous verify, highlight lifecycle, artifact no-auto-promotion, vision policy, screenshot scope/privacy, DPI/stale/destructive guards, custom-drawn fixture path, double-click policy, STJ `windowHandle` hex serialization. `/debug/status` covered in `PublicContractProtectionTests`.

### Windows acceptance

**OVERALL PASS (9/9)** on tip `28b8712` — evidence on #148 under `docs/acceptance/phase8-*`. `/debug/inspect` HTTP re-smoke **OK** on `65a3326`. Duplicate PR #149 closed; #148 owns evidence.

See [`docs/acceptance/phase8-windows-acceptance.md`](acceptance/phase8-windows-acceptance.md).

### Publish results

**CI green** on tip `65a3326` (4/4 — Build and Test + Publish Standalone Packages).

### git diff --check

Clean against `origin/main...HEAD` at `65a3326`.

### Known limitations

- Real OCR providers not enabled (fake only)
- No headless claim; interactive desktop required
- Vision independently gated; Debug Mode does not enable vision

### Exit gate

Debug capture and visual fallback must work on controlled fixtures without intercepting clicks, persisting sensitive screenshots by default, or promoting any candidate automatically. **Met** for review: CI green on `65a3326`, Windows 9/9 PASS, inspect re-smoke OK. PR stays draft; do not merge unless instructed.
