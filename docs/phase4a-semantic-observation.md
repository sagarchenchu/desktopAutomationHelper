# Phase 4A — Semantic desktop observation

Observe-only semantic snapshots of the current desktop application window for
explorers, planners, and later MCP clients. This phase does **not** click,
hover, type, crawl, plan, call Jira, or invoke AI.

## Contract

Schema: [`automation/schemas/semantic-snapshot.schema.json`](../automation/schemas/semantic-snapshot.schema.json)
(`snapshotSchemaVersion: 1`).

Checked-in example (no customer data/secrets):
[`automation/observations/examples/sample-semantic-snapshot.json`](../automation/observations/examples/sample-semantic-snapshot.json).

Required top-level fields:

- `snapshotSchemaVersion`, `snapshotId`, `capturedAtUtc`, `driverVersion`, `snapshotHash`
- `application` (`processId`, `processName`)
- `windowsSession` (`sessionId` only — no secrets)
- `window` (ephemeral `ref`, HWND diagnostic metadata, title, class, processId, modal, optional bounds)
- ordered `nodes` with ephemeral `ref` / `parentRef`, ordinal, depth, controlType/role,
  Name, AutomationId, ClassName, optional HelpText, enabled/offscreen/focus flags,
  UIA patterns, pattern-derived capabilities, optional bounds
- `truncated`, `warnings`, stable `ordering.rule`

### Ordering

Depth-first pre-order via parent tree links (not a global depth|ordinal sort).
Only siblings are sorted (ordinal, then `automationId`, `name`, `controlType`,
`nodeIndex`). Identical normalized trees produce the same `snapshotHash`
(hash ignores `snapshotId` / `capturedAtUtc` / runtime IDs / PID / session ID /
HWND / bounds).

### Ephemeral refs

- Opaque (`n0`, `n1`, …) and scoped to one snapshot.
- Never treat as persistent locators or approved knowledge.
- `EphemeralRefIndex` rejects unknown, duplicate, and cross-snapshot (stale) refs.
- Optional driver `runtimeId` may assist in-memory resolution and is **not** persisted.

## Privacy and size controls

- Default `actionableOnly=true` keeps content/action-relevant nodes (ancestors retained).
- Bounded `maxDepth`, `maxChildren`, `maxNodes` (default agent maxNodes = 500).
- Truncation is marked with warnings — content is never silently dropped past caps.
- Editable field values are omitted by default (never projected into the snapshot).
- `SecretRedactor` runs before artifact write and on bounded strings.
- Bounds/HelpText are opt-in (`--include-bounds`, `--include-help-text`).
- Response-size protection reuses runner `MaxResponseBytes`.

## Observation source

Reused and extended **`dumpuia`** (no new `/ui` route). Catalog schema remains **version 2**.

Optional dumpuia inputs (defaults preserve capture-page behavior):

| Input | Default | Purpose |
|---|---|---|
| `maxNodes` | hard cap 2000 when omitted | Soft truncation with warnings |
| `includeBounds` | `true` | Omit rectangles when `false` |
| `includeHelpText` | `false` | HelpText |
| `includeFocusState` | `false` | focusable / hasKeyboardFocus |
| `includeRuntimeId` | `false` | Ephemeral runtimeId |
| `includeTreeLinks` | `false` | nodeIndex / parentNodeIndex / ordinal |
| `includeWindowState` | `false` | Window-pattern modal/state |

Successful dumpuia payloads include `semanticDiagnosticVersion: 1`. Hitting
`maxDepth` / `maxChildren` / `maxNodes` marks `truncated` with warnings (no silent
drop). ValuePattern exposes `valueIsReadOnly`; `setValue` is emitted only when
writable (`CurrentIsReadOnly == false`) — fail closed for Phase 4B.

Agent `observe-app` requires `semanticDiagnosticVersion` and tree-link fields,
requests focus/window state and a bounded `maxNodes`, then normalizes into the
semantic snapshot DTOs. Older drivers that omit the marker/tree links fail fast
instead of producing flat degraded snapshots.

## CLI

```cmd
dotnet run --project src/DesktopAutomationAgent -- observe-app --json
dotnet run --project src/DesktopAutomationAgent -- observe-app --json ^
  --view control --root activeWindow --max-depth 8 --max-nodes 400 ^
  --include-bounds --output observations/manual/snapshot.json
```

- Structured JSON → stdout; logs → stderr.
- Default artifact: `automation/observations/<observationId>/snapshot.json`
- Explicit `--output` must stay inside the workspace and must not overwrite.

Requires an active driver session (for example `switchwindow` / `launch`) so
dumpuia can resolve the application window, matching `capture-page`.

## Agent surface

- DTOs under `DesktopAutomationAgent.Observation`
- `ISemanticObservationService` / `SemanticObservationService`
- `SemanticSnapshotNormalizer`, `EphemeralRefIndex`, `PatternCapabilityMapper`
- Exactly one diagnostic `POST /ui` (`dumpuia`) per observation — no action ops

## Exact driver diff (summary)

- `UiRequest`: optional semantic diagnostic fields listed above
- `NativeUiaTreeDiagnosticService`:
  - soft `maxNodes` truncation + `truncationWarnings`
  - optional HelpText, focus flags, runtimeId, tree links, window state
  - pattern → capability mapping; bounded strings
  - `rootDiagnostics` gains `processName` and `windowsSessionId`
- No new HTTP routes; `GET /ui/operations` schema version stays **2**

## Limitations

- Observe-only: no crawl, action-by-ref, knowledge graph, Jira/BDD, plan, MCP, AI, OCR, or Debug Mode.
- Active session / resolvable root still required for meaningful window scope.
- `maxDepth` / `maxChildren` mark informational `truncated` without stopping sibling
  traversal; only `maxNodes`, timeout, or cancellation set stop-traversal.

## Test evidence

- Agent: `Phase4aSemanticObservationTests` (DFS with grandchild, hash with differing
  bounds, maxNodes/truncation, capabilities, disabled/offscreen, value omission,
  redaction, stale/duplicate refs, single dumpuia POST, concurrent/non-overwrite
  artifact writes, oversized response rejection, schema/sample validation).
- Driver: `NativeUiaSemanticContractTests` bounded traversal contracts
  (maxChildren keeps first N siblings; maxDepth visits siblings and omits
  grandchildren; maxNodes stops traversal) plus dumpuia/catalog suites.
- Phase 0–3 suites remain green aside from known symlink path-safety tests that
  require `SeCreateSymbolicLinkPrivilege`.

### Windows acceptance (self-hosted worker)

Required before merge — see
[`docs/acceptance/phase4a-windows-acceptance.md`](acceptance/phase4a-windows-acceptance.md).

- **WPF** `tools/DesktopAutomationWpfSmokeApp` (`DA WPF Smoke`): observed menu,
  button, textbox, combo, list, disabled control, and child dialog under correct
  process/session. Harness PR #138 alone is not enough for this gate.
- **WinForms** `DesktopAutomationTestApp` (`DA Test Harness`, harness PR #138 when
  not on main): observed menu, button, textbox, combo, DataGrid/`SampleGrid`,
  disabled control, child window, and **actual modal** `DA Sample Dialog`.
- EventLog `getvalue` unchanged across observe-only (no UI mutation from observation).
