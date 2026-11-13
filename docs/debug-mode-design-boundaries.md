# Debug Mode design boundaries

Debug Mode is implemented in Phase 8 as isolated services. Vision fallback is a
related but independently gated deliverable.

## Do not grow the overlay or UiService for Debug Mode

Ctrl+D / Debug Mode feature logic must **not** be placed directly into the large
`RecordingOverlayWindow` or `UiService` classes.

The overlay remains a **UI/menu coordinator only** (thin Ctrl+D toggle and
status text forwarding to `IDebugSessionCoordinator`).

## Service split (actual Phase 8 names)

| Service | Namespace | Role |
|---|---|---|
| `DebugSessionCoordinator` | `DesktopAutomationDriver.Services.Debug` | Session start/stop, inspect orchestration |
| `UiaTreeCaptureService` | `DesktopAutomationDriver.Services.Debug` | Bounded UIA ancestry/children/patterns |
| `LocatorCandidateGenerator` | `DesktopAutomationDriver.Services.Debug` | Object-repository / Phase 3 candidate locators |
| `NonInteractiveHighlightService` | `DesktopAutomationDriver.Services.Debug` | Yellow non-interactive highlight |
| `CandidateVerificationService` | `DesktopAutomationDriver.Services.Debug` | Unique / ambiguous verification |
| `DebugArtifactWriter` | `DesktopAutomationDriver.Services.Debug` | Candidate artifacts only |

Dependency direction:

```text
RecordingOverlayWindow  -->  IDebugSessionCoordinator
DebugSessionCoordinator -->  UiaTreeCaptureService
                            LocatorCandidateGenerator
                            NonInteractiveHighlightService
                            CandidateVerificationService
                            DebugArtifactWriter
                            IUiaElementProbe
```

Vision services (`VisionFallbackPolicy`, `VisionScreenshotService`,
`IVisionProvider` / `FakeVisionProvider`, `CoordinateActionGuard`,
`VisionFallbackCoordinator`) live under
`DesktopAutomationDriver.Services.Vision` and must not be folded into
`UiService` or the overlay.

## Hard rules

- No automatic promotion of candidates
- Vision disabled by default; never replaces working UIA locators
- Highlight must not intercept clicks
- Interactive desktop required for coordinate / hover / screenshot workflows
  (no headless claims in this phase)

See also: [`phase8-debug-vision.md`](phase8-debug-vision.md).
