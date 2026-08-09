# Plans

Reusable deterministic command plans for the DesktopAutomationAgent runner.

## Layout

- `example.plan.json` — session-free smoke example that calls `listwindows` only.
- `../schemas/plan.schema.json` — JSON Schema (Draft 2020-12) for offline validation.

## Authoring rules

- `schemaVersion` must be `1` and `catalogSchemaVersion` must be `2`.
- `planId` must match `^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$`.
- `steps` is required and must contain at least one step. `onFailureSteps` and `cleanupSteps` are optional.
- Combined `steps` + `onFailureSteps` count must not exceed `1000`.
- Step `id` values must be unique (case-insensitive) across all step lists.
- Each step requires `operation` (no leading/trailing whitespace) and an `arguments` object.
- Do not place `operation`, `authorization`, or `bearerToken` inside `arguments`.
- Cleanup steps must not define `assertions` or set `captureResponse`.
- Plans that call `launch` must end with `close` or `quit`, and `onFailureSteps` must include `close` or `quit`. `closewindow` does not end a session.
- Do not store secrets in plan files. Mark sensitive argument values with `sensitive: true` on the step when needed.

Validate a plan offline with the agent runner (Phase 2) or against `plan.schema.json` in your editor.
