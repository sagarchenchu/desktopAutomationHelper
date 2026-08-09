# Phase 0 — Desktop Automation Driver Baseline

## Baseline identity

| Item | Value |
|------|-------|
| Repository | `https://github.com/sagarchenchu/desktopAutomationHelper` |
| Baseline commit SHA | `c4f46145c3ea81565395620a3e8b877c2e42fb4c` |
| Driver version | `1.0.105` (`DesktopAutomationDriver.csproj` `<Version>`) |
| Working branch at start | `copilot/rewrite-combobox-select-native-uia-only` (clean tree at baseline SHA) |

## Build / test commands

```bash
dotnet build DesktopAutomationHelper.slnx --configuration Release
dotnet test src/DesktopAutomationDriver.Tests --configuration Release
```

**Environment note (this agent host):** the checked-in `.slnx` format is not recognized by .NET SDK 8.0.423 MSBuild on Linux (`MSB4068`). Project-level build was used instead:

```bash
dotnet build src/DesktopAutomationDriver/DesktopAutomationDriver.csproj --configuration Release
dotnet test src/DesktopAutomationDriver.Tests/DesktopAutomationDriver.Tests.csproj --configuration Release
```

**Baseline build result:** succeeded (0 errors; existing nullable warnings only).

**Baseline test result:** aborted before any tests executed — `Microsoft.WindowsDesktop.App` 8.0 runtime is not available on this Linux host. No passed/failed/skipped counts were produced. This is an environment limitation, not an assertion failure in the suite. CI on `windows-latest` remains the authoritative test runner.

## Post-implementation verification (this host)

| Check | Result |
|-------|--------|
| `dotnet build ...DesktopAutomationDriver.csproj -c Release` | Succeeded |
| `dotnet build ...DesktopAutomationDriver.Tests.csproj -c Release` | Succeeded (0 errors) |
| `dotnet test ... -c Release` | Aborted (missing WindowsDesktop runtime); Windows CI is authoritative |
| Catalog ↔ `UiService` switch parity (static) | 137 / 137 names covered |

## Follow-up corrections (PR #128 review)

- Fixed three playback contract tests (`Click Submit`, supported `hover` ActionType, `continueOnError:false` action count = 1).
- Catalog `schemaVersion` = 2 with `requiredInputAlternatives` for multi-shape inputs.
- `IsKnownOperation` no longer trims (matches `UiService.Execute`).
- Audited metadata for launch, switchwindow, select, scroll, dragbyoffset, getposition, contextmenupath/tree paths, sendkeysuia, inspectcombobox, iseditable, finduia.
- `dumpuia` and all `popup-alert` ops: `requiresSession: false` (processId / desktop discovery).
- `findlocator` alternatives: locator | locatorPath | criteria; `finduia` alternatives: locator | nameContains | bestMatch | hwnd | className.
- Session-optional coordinate/global paths: `closewindow`, `listtrackedwindows`, `scroll`, `mousescroll`, `dragbyoffset`, `dragcoordinates`, `mouse`, `sendkeysuia`.
- `clickmenu` requires `value`; `popupok` alternatives: value | hwnd | className.
- value/index alternatives: `selectcomboboxuia`, `selectopendropdownitem`, `clickheaderdropdownitem`, `selectheaderdropdownitem`.
- `mousescroll`: locator | x+y; `mouse`: action+x+y | action+fromX+fromY+toX+toY.
- `UiOperationCatalogResponse.SchemaVersion` default aligned to 2.

## Public API compatibility statement

Phase 0 is a stabilization phase. It:

- Adds `GET /ui/operations` discovery without changing existing `/ui` dispatch behavior.
- Adds contract/parity tests for recording, playback, diagnostics, auth and JSON shapes.
- Does **not** rename routes, request/response properties, operation aliases, timeout defaults, locator resolution, recorder export format, playback mapping, FlaUI/Native UIA routing, bearer auth, `/verify`, port calculation, or failure screenshots.
- Does **not** bump the driver version.

## Known incomplete / experimental components

- Release workflow tags use `v1.0.${{ github.run_number }}`, which can diverge from `<Version>` in the `.csproj`.
- Some Native UIA operations remain best-effort against complex enterprise controls.
- Assistive recording overlay requires an interactive Windows desktop (not covered by unit tests).

## Phase 0 scope confirmation

No Jira integration, Gherkin/BDD compilation, AI/LLM/Copilot integration, database connectivity, suite orchestration, agent projects, or self-healing behavior was added in Phase 0.
