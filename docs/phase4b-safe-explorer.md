# Phase 4B — Safe navigation explorer

Policy-bounded desktop explorer that repeatedly observes, selects one explicitly
permitted navigation action, executes it through `IDriverUiClient`, waits for
stabilization, observes again, and records a candidate transition.

**No AI.** Modes are `observe` and `navigate` only. Deny-by-default.

## Services

| Service | Role |
|---|---|
| `ExplorationCoordinator` | Bounded explore loop (not in AgentCli / DriverUiClient / DeterministicPlanRunner) |
| `ExplorationPolicyReader` | Load + validate policy JSON |
| `CapabilityResolver` | Snapshot patterns → technical candidates; authorize via policy |
| `FrontierQueue` | Deterministic frontier; never repeats a state/action pair |
| `StateFingerprinter` | Stable state identity (ignores refs/HWND/PID/bounds/timestamps) |
| `UiStabilizationService` | Poll until quiet period or timeout |
| `WindowTransitionDetector` | New/closed/modal/external window + scope checks |
| `TransitionRecorder` | In-memory states/transitions |
| `ExplorationArtifactWriter` | Atomic, non-overwriting, redacted artifacts |

## Modes

1. **observe** — snapshot + seed candidates only; no UI mutation
2. **navigate** — execute only policy-approved navigation actions

Not implemented: test-data, transactional, destructive, unrestricted.

## Policy contract

- Schema: [`automation/schemas/exploration-policy.schema.json`](../automation/schemas/exploration-policy.schema.json)
- Example: [`automation/exploration/policies/example.policy.json`](../automation/exploration/policies/example.policy.json)

Required fields include application identity, process/session/window scope,
allowed action categories, explicit allow/deny rules, budgets (states, actions,
depth, duration, per-action timeout, stabilization),
`allowCloseChildWindowsForBacktrack`, `allowHover`, and
`allowCoordinateFallback` (must be `false` in Phase 4B).

## Deny by default

- `InvokePattern` / click capability alone never authorizes an action.
- Labels matching delete/remove/approve/submit/save/pay/purchase/send/transfer/publish/install/uninstall are always denied in Phase 4B.
- String heuristics are an additional guard, not the sole authorization rule.
- Cross-process windows are rejected unless listed in `allowedChildProcessNames`.
- Coordinate fallback and double-click are rejected by policy validation.

## Exact allowed action kinds (navigate)

When authorized by category and/or explicit allow rules:

| Kind | Typical driver op | Notes |
|---|---|---|
| Expand / Collapse | `expandtreeitem` / `collapsetreeitem` | Prefer over invoke |
| TabSelect | `selecttabuia` | TabItem |
| MenuOpen | `clickuia` | Expandable Menu / MenuItem containers only (leaf Exit/Reset/Print/Archive → Invoke + explicit allow) |
| SelectItem | `select` | List/tree selection items |
| Scroll | `scroll` (pattern mode) | When ScrollPattern present |
| Focus | `focusuia` | When focusable |
| Hover | `hover` | Only when `allowHover=true` |
| Invoke | `clickuia` | **Requires explicit allow rule** |

Never: arbitrary typing, save/submit/delete, double-click, coordinate clicks.

## State algorithm

1. Capture before snapshot
2. Fingerprint
3. Select one allowed frontier action deterministically
4. Record intended action
5. Execute through `IDriverUiClient`
6. Stabilize (bounded quiet period)
7. Capture after snapshot
8. Diff fingerprints / window inventory
9. Record success, no-op, failure, modal, new/closed window, or scope violation
10. Enqueue only previously unseen permitted frontier actions

Cancellation always attempts a valid partial artifact write.

## State identity risks

Fingerprint ignores volatile timestamps, ephemeral refs, handles, coordinates,
PID, and session id. It includes process name, window title/class/modal, visible
control identity/enabled/offscreen/patterns/capabilities, and interaction state
(selected/expanded/toggle) when dumpuia projects it.

**Collision risk:** expand/selected-only mutations collide only when interaction
state is unavailable from the driver. **False-difference risk:** truncation or
focus-adjacent name changes can create extra states.

## Budgets and cycle protection

- Never execute the same state fingerprint + action key twice in one crawl
- Enforce maxStates, maxActions, maxDepth, totalDuration, per-action timeout,
  stabilization timeout even while the UI keeps changing
- No unbounded recursion (iterative frontier)

## CLI

```cmd
dotnet run --project src/DesktopAutomationAgent -- explore-app --policy automation/exploration/policies/example.policy.json --dry-run --json
dotnet run --project src/DesktopAutomationAgent -- explore-app --policy automation/exploration/policies/example.policy.json --json
```

Dry-run lists permitted candidates and performs no action `POST /ui`.

## Artifacts

```
automation/exploration/runs/<explorationId>/
  manifest.json
  states.jsonl
  transitions.jsonl
  summary.json
```

Atomically written, non-overwriting, redacted, gitignored by default.

## Windows test fixture

`tools/DesktopAutomationExplorerFixture` — WinForms app titled
**DA Explorer Fixture** with menu, tabs, expandable tree, hover target, list,
scroll panel, modal dialog, second window, and clearly marked destructive
controls (`Delete`, `Submit`, `Save`, `Purchase`) that the crawler must refuse.

## Safety limitations

- No recovery/self-healing
- No vision/OCR
- No persistent approved knowledge
- No transactional/test-data modes
- Selected/expanded pattern state is not yet a first-class snapshot field
