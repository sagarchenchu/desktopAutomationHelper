# Phase 4 — Deterministic BDD Plan Compiler

Phase 4 adds an **offline** Agent CLI that compiles Assistive Mode artifacts into
Phase 2 execution plans.

It does **not** call Jira, interpret Gherkin/NLP, invent locators, or use AI.

## Goal

Same local inputs → same plan bytes.

Inputs:

1. `bdd-action-map.json` (structure, operations, object refs, BDD group ids)
2. Source recording JSON (values, element locators, menu paths, metadata)
3. Optional promoted object repository when `--prefer-object-refs` is set

Output:

- A Phase 2 `*.plan.json` under `automation/plans/` (or `--out`)

## CLI

```cmd
dotnet run --project src/DesktopAutomationAgent -- compile-bdd ^
  --map path/to/bdd-action-map.json ^
  [--recording path/to/recording_*.json] ^
  [--out plans/my.plan.json] ^
  [--prefer-object-refs --object-repository object-repository/repository.json] ^
  [--json]
```

- `--recording` may be omitted when `sourceRecording` can be resolved by walking up
  from the map directory (Assistive layout: recording at output root, map under
  `assistive-artifacts/<jira>/<recordingId>/`).
- Default output: `plans/<jiraKey.lower>.<recordingId>.plan.json`
- Default locator strategy: **inline locators from the recording** (candidates are
  not auto-promoted). `$objectRef` emission requires `--prefer-object-refs` and a
  repository path after manual promotion.

Exit code `5` on compilation/validation failure (same family as suite/plan validation).

## Deterministic rules

1. BDD `statement` text is copied into plan metadata only. It is never parsed into operations.
2. Operations come from the map’s `operation` field (already resolved by Assistive export).
3. Action order follows map `bddGroups` order, then each group’s `sequence` / `eventId`.
4. Values and locators are joined from the recording by `eventId`.
5. `unmappedEventIds` are reported and not compiled.
6. The emitted plan is validated with the existing `PlanValidator` before write.
7. No HTTP, Jira, AI, or driver project reference.

## What this is not

| Capability | Status |
|---|---|
| Jira REST / auth | Out of scope |
| Gherkin feature-file parsing | Out of scope |
| Suite execution / scheduling | Out of scope |
| Auto-promote candidate pages | Out of scope |
| Self-healing / AI repair | Out of scope |
| Debug Mode | Out of scope |

## Workflow

1. Record Assistive session with Jira/BDD annotations (driver).
2. Review candidate page objects; promote manually if desired (Phase 3).
3. `compile-bdd --map ...` → plan file.
4. `validate-plan --file ...`
5. `run-plan --file ...` (or `--dry-run`)

## Related docs

- [Assistive BDD recording](assistive-bdd-recording.md)
- [Phase 2 deterministic runner](phase2-deterministic-runner.md)
- [Phase 3 object repository](phase3-object-repository.md)
- [Phase 3.5 hardening](phase3-5-post-merge-hardening.md)
