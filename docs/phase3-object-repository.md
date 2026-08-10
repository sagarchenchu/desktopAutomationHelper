# Phase 3 — Object Repository

Phase 3 adds an offline-validated object repository, capture and verify tooling,
plan `$objectRef` expansion, and CLI commands. The agent still communicates with
the driver only through existing HTTP (`IDriverUiClient` / `IDriverCatalogClient`).

See also: [Phase 2 deterministic runner](phase2-deterministic-runner.md) for plan
execution and `$objectRef` usage in plans.

## Layout

```text
automation/
  object-repository/
    repository.json       # manifest (tracked)
    pages/                # active page-object documents (tracked)
    candidates/           # generated drafts (gitignored)
    captures/             # raw dumpuia output (gitignored)
  schemas/
    object-repository.schema.json
    page-object.schema.json
```

## Identifiers and references

- `repositoryId`, `pageId`, and element keys: `^[a-z][a-z0-9-]{0,63}$`
- Object references: `pageId.elementId` (for example `login.submit-button`)

## Locator rules

Allowed locator fields:

- `automationId` (alone is sufficient)
- `name` + `controlType`
- `className` + `controlType`
- `matchMode`: `exact`, `contains`, `startswith`
- `foundIndex` (discouraged; fragile warnings)

Volatile fields (handles, coordinates, runtime IDs, bounding boxes, etc.) are rejected.
Blank/whitespace string values are rejected when a property is present.

Quality grades: `strong`, `medium`, `weak` (capture never auto-assigns `weak`).

## Manifest validation

`ObjectRepositoryValidator` enforces:

- Manifest-referenced pages must have `state: "active"`
- Page files must be under `pages/` (no `..`)
- Duplicate `pageId` or duplicate page file paths are errors
- Non-preferred file names under `pages/` produce warnings (`pages/<pageId>.page.json` preferred)
- Active pages must not contain a non-empty `unresolved` section
- Active pages must not contain `source.kind: "capture"` elements

## Plan `$objectRef` expansion

Plans may declare:

```json
"objectRepository": "object-repository/repository.json"
```

Steps may reference repository objects only in these argument keys:

- `locator`, `locator2`, `parentLocator`, `containerLocator`

Marker shape (exactly one property):

```json
{ "$objectRef": "login.submit" }
```

Expansion runs offline after plan structure validation and before driver catalog
preflight. Missing/invalid references fail validation (exit `5`). Expanded locators
are sent in `POST /ui` bodies; `$objectRef` never reaches the driver.

## Capture workflow (`capture-page`)

1. Validate repository manifest (empty `pages` is OK)
2. Driver `/status`, `/ui/operations`, catalog compatibility
3. Confirm `dumpuia` is canonical, non-deprecated, `operationType: diagnostic`
4. Exactly one `POST /ui` with `dumpuia`
5. Write atomically (no overwrite):
   - `captures/<pageId>/<captureId>.capture.json`
   - `candidates/<pageId>/<captureId>.page.json`

Default capture arguments: `view=control`, `root=activeWindow`, `maxDepth=8`,
`maxChildren=200`, `includeOffscreen=false`, `includePath=true`,
`timeoutMs` from `ObjectRepository:DiagnosticTimeoutMilliseconds`.

## Verify workflow (`verify-object-repository`)

1. Offline repository validation
2. Driver resolve, `/status`, catalog, confirm `finduia` diagnostic operation
3. Select objects: all active pages, `--page`, or `--ref` (mutually exclusive)
4. Sequential one `finduia` per object, no retries
5. Match rules: 0 missing, 1 pass, >1 ambiguous unless `foundIndex` in range (fragile warning)

## CLI commands

```bash
dotnet run --project src/DesktopAutomationAgent -- validate-object-repository \
  --file automation/object-repository/repository.json

dotnet run --project src/DesktopAutomationAgent -- resolve-object \
  --file automation/object-repository/repository.json --ref login.submit

dotnet run --project src/DesktopAutomationAgent -- capture-page \
  --file automation/object-repository/repository.json --page login --name "Login"

dotnet run --project src/DesktopAutomationAgent -- verify-object-repository \
  --file automation/object-repository/repository.json --page login

dotnet run --project src/DesktopAutomationAgent -- validate-plan \
  --file automation/plans/example.plan.json
```

Offline commands (`validate-object-repository`, `resolve-object`, `validate-plan`
expansion) make zero HTTP calls.

## Configuration

`automation/config/agentsettings.example.json`:

```json
"ObjectRepository": {
  "MaxFileBytes": 5242880,
  "MaxPages": 500,
  "MaxElementsPerPage": 5000,
  "MaxTotalElements": 50000,
  "DiagnosticTimeoutMilliseconds": 15000
}
```

CLI overrides: `--ObjectRepository:...` or `ObjectRepository__...`.

## Exit codes

| Code | Meaning |
| --- | --- |
| 0 | Success |
| 2 | Usage / configuration |
| 3 | Driver unavailable |
| 4 | Auth / catalog |
| 5 | Suite, plan, workspace, object repository validation |
| 6 | Capture, verify, UI operation, timeout, assertion failure |
| 7 | Cancelled |

## No AI / no driver coupling

Phase 3 does not call AI providers or reference `DesktopAutomationDriver` from the
agent project. Capture and verify use `dumpuia` / `finduia` over HTTP only.

## Promotion workflow

1. `capture-page` writes capture + candidate artifacts
2. Human review; edit locators; remove PII from `name` values
3. Copy approved page JSON to `pages/` with `state: "active"` and `source.kind: manual` or `approved`
4. Add page entry to `repository.json`
5. Commit manifest and `pages/` only

## Run report fields

When a plan uses the object repository, `run.json` may include:

- `objectRepositoryPath`
- `objectRepositoryId`
- `objectRepositorySha256`
- `resolvedObjectReferences` (sorted, distinct)
