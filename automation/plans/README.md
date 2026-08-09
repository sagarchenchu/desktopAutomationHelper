# Plans

Plans are compiled executable artifacts for the Desktop Automation Agent
deterministic runner (Phase 2).

## Layout

- `example.plan.json` — session-free smoke example that calls `listwindows` only.
  The driver returns a root JSON array of window descriptors; assert `path: ""` / `isNotNull`.
  Do not pass `limit` unless/until the driver supports it.
- `../schemas/plan.schema.json` — JSON Schema (Draft 2020-12) for offline validation.

## Phase 2 authoring

- Phase 2 supports **manually authored** plans.
- Later phases will compile Jira BDD into the same format.
- Existing valid plans execute without AI, Jira, object-repository, or database access.
- `schemaVersion` must be `1` and `catalogSchemaVersion` must be `2`.
- `planId` must match `^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$`.
- `steps` is required and must contain at least one step. `onFailureSteps` is optional.
- Combined `steps` + `onFailureSteps` count must not exceed `1000`
  (authoritative rule enforced by the C# `PlanValidator`; schema `maxItems` are editor hints only).
- Step `id` values must be unique (case-insensitive) across all step lists.
- Each step requires `operation` (no leading/trailing whitespace) and an `arguments` object.
- Do not place `operation`, `authorization`, or `bearerToken` inside `arguments`.
- Cleanup steps must not define `assertions` or `captureResponse`.
- Plans that call `launch` must end with `close` or `quit`, and `onFailureSteps` must include `close` or `quit`. `closewindow` does not end a session.
- Plans must not contain credentials. Password-entry steps must use `sensitive: true`.

```bash
dotnet run --project src/DesktopAutomationAgent -- validate-plan --file automation/plans/example.plan.json
dotnet run --project src/DesktopAutomationAgent -- run-plan --file automation/plans/example.plan.json --dry-run
```
