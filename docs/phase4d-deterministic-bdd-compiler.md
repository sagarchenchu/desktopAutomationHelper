# Phase 4D — Deterministic BDD compiler and plan cache

Phase 4D compiles **offline** Jira-exported BDD into the existing Phase 2 plan
format. It does **not** call Jira, invent steps with AI, crawl the UI, or
`POST /ui`.

## Goal

Identical semantic inputs → identical executable `plan.json` bytes.

Inputs:

1. Offline Jira test-case export (`automation/schemas/jira-test-case.schema.json`)
2. Approved step registry (`automation/step-registry/registry.json`)
3. Approved object repository (for `$objectRef` resolution)
4. Optional approved application knowledge (navigation refs only)
5. Optional operations catalog snapshot (catalog preflight only)

Output cache:

```
automation/compiled/<jiraKey>/<fingerprint>/
  plan.json
  compilation.json
```

## Dependency / license

| Component | Package | License | Version pin |
|---|---|---|---|
| Gherkin parser | NuGet `Gherkin` (Cucumber) | MIT | `30.1.2` (`gherkin-30.1.2`) |

Compiler contract version: `phase4d-1`.

Phase 4D uses the official Cucumber Gherkin parser for Feature / Background /
Scenario / Given / When / Then / And / But. It does **not** claim full Gherkin:
data tables, doc strings, Scenario Outline / Examples, and Rule are rejected
with precise `unsupported-feature` errors unless/until implemented and tested.

## BDD v1

Supported:

- `Feature`, `Background`, `Scenario`
- `Given` / `When` / `Then` / `And` / `But`
- Quoted parameters via step-registry exact/`{param}` patterns or named regex groups
- Structured JSON scenarios (compiled to Gherkin text, then the same IR path)

Unsupported (fail closed):

- Data tables
- Doc strings
- Scenario Outline / Examples
- Rule
- Credentials / raw production secrets in exports

`And` / `But` normalize to the preceding primary keyword.

## Step registry

Schema: `automation/schemas/step-registry.schema.json`

Each mapping declares:

- `mappingId`
- exact or bounded-regex `pattern` (culture-invariant, timeout-bounded, compiled once; regex must be `^…$`)
- keyword restrictions (`Given` / `When` / `Then`)
- named parameters + types (`string` / `int` / `bool` / `objectRef`)
- Phase 2 plan-step templates
- approved object / knowledge refs
- expected state / postcondition
- `risk` (`low` / `medium` / `high` / `prohibited`)
- source / version / approval

Match rules:

- 0 matches → `unmapped-step`
- >1 matches → `ambiguous-step`
- `risk: prohibited` → `prohibited-risk`
- Object refs resolve against **approved/active** repository pages only
- Knowledge refs require **approved** knowledge; missing / conflicting / ambiguous fail closed

## Compilation pipeline

1. Validate export (schema + canonical Jira key; no network)
2. Parse → immutable IR
3. Normalize And/But
4. Registry match
5. Resolve approved objects / knowledge
6. Emit Phase 2 `PlanManifest`
7. Run existing `PlanValidator`
8. Object-reference expansion check (plan retains `$objectRef`; expansion must succeed)
9. Optional catalog preflight when `--catalog` is supplied
10. Write cache atomically; never overwrite a different artifact; never `POST /ui`

## Fingerprint

SHA-256 over a canonical object of:

| Field | Notes |
|---|---|
| `inputSha256` | BDD / export bytes |
| `jiraRevision` | `revision` or `updatedAt` |
| `compilerVersion` | `phase4d-1` |
| `parserVersion` | `gherkin-30.1.2` |
| `stepRegistrySha256` | registry file hash |
| `objectRepositorySha256` | repository aggregate hash |
| `knowledgeSha256` | included **only when** knowledge was used |
| `catalogFingerprint` | catalog schema + sorted operation-set hash when catalog requested |

Reuse only on full fingerprint match with identical plan bytes.

Determinism rules:

- Timestamps appear only in `compilation.json` metadata
- Executable plan bytes use canonical JSON (sorted properties, stable ordering)
- Scenario / tag / mapping ordering is stable

## CLI

```cmd
dotnet run --project src/DesktopAutomationAgent -- validate-test-case --file automation/test-cases/examples/DEMO-1.json
dotnet run --project src/DesktopAutomationAgent -- validate-step-registry
dotnet run --project src/DesktopAutomationAgent -- compile-key --key DEMO-1 --explain --json
dotnet run --project src/DesktopAutomationAgent -- validate-compiled-key --key DEMO-1 --json
```

`compile-key` never calls Jira, AI, or `POST /ui`.

## Related docs

- [Phase 2 deterministic runner](phase2-deterministic-runner.md)
- [Phase 3 object repository](phase3-object-repository.md)
- [Phase 4C application knowledge](phase4c-application-knowledge.md)
