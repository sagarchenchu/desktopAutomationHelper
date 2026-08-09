# Phase 2 — Deterministic Plan Runner

Phase 2 adds an offline-validated, fail-fast runner that executes already-compiled
JSON command plans against `DesktopAutomationDriver` over HTTP.

> Phase 2 performs no Jira, BDD, AI, object-repository, database, scheduling or
> suite orchestration work.

Phase 2 **executes plans**. It does not understand BDD and does not generate or
repair plans. A future compiler may emit the same plan format from Jira BDD;
once generated, repeated executions must not require AI calls.

## HTTP boundary

The agent may call only:

| Method | Endpoint | Purpose |
| --- | --- | --- |
| `GET` | `/verify` | Existing safe discovery (Phase 1) |
| `GET` | `/status` | Driver readiness |
| `GET` | `/ui/operations` | Operation catalog |
| `POST` | `/ui` | Execute one deterministic plan step |

`validate-plan` makes **no** HTTP calls.  
`run-plan --dry-run` may call `/verify`, `/status`, and `/ui/operations`, but never `POST /ui`.  
Only a real `run-plan` may call `POST /ui`.

## Deterministic guarantees

1. The plan is read once and hashed (SHA-256 of exact bytes).
2. Every main and cleanup step is validated before the first `POST /ui`.
3. Main steps execute in declared array order.
4. Steps are never reordered, inserted, removed, rewritten, or retried by the agent.
5. Canonical operation names are required; aliases are rejected.
6. Execution stops after the first failed main step.
7. Cleanup (`onFailureSteps`) runs only after failure or cancellation, each step at most once.
8. Driver waits/retries remain controlled by the driver.

## Plan schema

Checked-in schema: `automation/schemas/plan.schema.json` (JSON Schema Draft 2020-12).

Required top-level fields: `schemaVersion` (`1`), `catalogSchemaVersion` (`2`),
`planId`, `name`, `steps`.

Optional: `$schema`, `description`, `tags`, `metadata`, `onFailureSteps`.

Example (session-free): `automation/plans/example.plan.json`.

When sending `POST /ui`, arguments are flattened:

```json
{
  "operation": "clickuia",
  "locator": { "automationId": "SubmitButton" },
  "timeoutMs": 20000
}
```

Do not nest arguments under an `arguments` property in the HTTP body.

## Main versus cleanup steps

- Main steps may include `assertions` and `captureResponse`.
- Cleanup steps (`onFailureSteps`) support `id`, `operation`, `arguments`, and
  `sensitive` only.
- Plans that `launch` must end with `close` or `quit`, and must include
  `close`/`quit` in `onFailureSteps`. `closewindow` does not end a session.

## Assertions

Assertions use RFC 6901 JSON Pointers against the driver response `value`.

Operators: `equals`, `notEquals`, `contains`, `matchesRegex`, `isTrue`,
`isFalse`, `isNull`, `isNotNull`.

They run only after a successful driver operation. All assertions on a step are
evaluated; any failure fails the step.

## Validation layers

### Offline (`validate-plan`)

Safe path, size, JSON, duplicates, schema versions, IDs, reserved argument names,
assertion shape. Does **not** claim operations are supported (catalog required).

### Catalog-aware (`run-plan` / `--dry-run`)

Resolve driver → `/status` (`ready=true`) → `/ui/operations` →
`CatalogCompatibility.Validate` → plan `catalogSchemaVersion` match →
canonical/alias checks → required inputs → session lifecycle.

## CLI examples

```bash
dotnet run --project src/DesktopAutomationAgent -- validate-plan --file automation/plans/example.plan.json
dotnet run --project src/DesktopAutomationAgent -- validate-plan --file automation/plans/example.plan.json --json
dotnet run --project src/DesktopAutomationAgent -- run-plan --file automation/plans/example.plan.json --dry-run
dotnet run --project src/DesktopAutomationAgent -- run-plan --file automation/plans/example.plan.json
dotnet run --project src/DesktopAutomationAgent -- run-plan --file automation/plans/example.plan.json --json
```

`--json` emits exactly one JSON document on stdout; logs go to stderr.

## Exit codes

| Code | Meaning |
| ---: | --- |
| 0 | Success or successful dry run |
| 2 | Usage or configuration error |
| 3 | Driver unavailable or unsafe discovery |
| 4 | Authentication or catalog incompatibility |
| 5 | Suite, plan, workspace or artifact validation failure |
| 6 | UI operation, timeout or assertion failure |
| 7 | Execution cancelled |

Unsupported operations and missing required inputs are code `5`.  
An invalid driver catalog itself is code `4`.  
Cleanup failure does not replace an established primary exit code.  
Failure to persist `run.json` is code `5` and never reports success.

## Run reports

Written atomically to `automation/runs/<runId>/run.json`.

Run IDs look like `20260809T203045123Z-a1b2c3d4`. Existing run directories are
never reused or overwritten.

Reports include plan identity/hash, status, exit code, dry-run flag, timestamps,
driver metadata, ordered step/cleanup results, failure classification, and
artifact-write outcome.

## Redaction

Bearer tokens and authorization material are never stored. Properties named like
`password`, `token`, `secret`, `apiKey`, `connectionString`, etc. are redacted
recursively. Steps with `sensitive=true` omit argument values, response values,
and assertion expected/actual values. `captureResponse=false` omits the response
`value`. Screenshot paths from the driver may be recorded; files are not copied
in Phase 2.

## Configuration (`Runner`)

```json
{
  "Runner": {
    "StepTransportTimeoutSeconds": 60,
    "CleanupTimeoutSeconds": 15,
    "MaxPlanBytes": 1048576,
    "MaxResponseBytes": 10485760,
    "RegexTimeoutMilliseconds": 500
  }
}
```

CLI/env forms: `--Runner:StepTransportTimeoutSeconds=60`,
`Runner__StepTransportTimeoutSeconds=60`, `DA_AGENT__Runner__...`.

Precedence remains: `appsettings.json` → `automation/config/agentsettings.local.json`
→ `DA_AGENT__*` → command-line.

## Known limitations

- No Jira/BDD compilation, AI repair, object repository, DB checks, suites,
  scheduling, notifications, parallelism, agent-level retries, templating, or
  HTML reporting.
- Agent remains `net8.0` and talks to the driver only through HTTP.
- Screenshot files are not copied into the run directory.

## Future connection to Jira/BDD

A later phase may compile Jira BDD into this same plan format. Phase 2 remains
the deterministic execution engine for those compiled artifacts.
