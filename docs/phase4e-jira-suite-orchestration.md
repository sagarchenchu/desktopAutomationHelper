# Phase 4E — Jira test-source and suite orchestration

Phase 4E adds a **provider-neutral, read-only** Jira test-case source and
sequential suite execution on top of the Phase 4D compiler and Phase 2 runner.

## Jira TM source identification

Before coding, this environment was inspected for a live test-management vendor:

| Probe | Result |
|---|---|
| Repo config / docs | No Xray, Zephyr, TestRail, or other TM plugin configured |
| Local `agentsettings.local.json` | Driver-only; no Jira vendor fields |
| Env (`JIRA_*` / `XRAY_*` / `ZEPHYR_*`) | None set |
| Existing artifact contract | Phase 4D normalized schema (`jira-test-case.schema.json`) |
| Assistive BDD recording | Local annotations only; no Jira API |

**Conclusion:** there is no assumed vendor model. Phase 4E exposes
`IJiraTestCaseSource` and maps vendor/plugin JSON through **configurable field
paths**. Mapping profiles (`plain-fields` default, plus `xray` / `zephyr` /
`normalized` / `custom`) only seed defaults — instance-specific custom field
IDs must be set in config.

## Goals

- Fetch only fields needed for the Phase 4D input model
- Cache redacted normalized artifacts safely
- Compile deterministically (Phase 4D)
- Execute a Jira key or suite via the existing Phase 2 runner
- Sequential only; fail-fast by default

## Auth / secrets

Secrets may come from:

1. Env (`DA_AGENT__Jira__ApiToken`, `DA_AGENT__Jira__BearerToken`, …)
2. Gitignored local config (`automation/config/agentsettings.local.json`)
3. Optional Windows Credential Manager target (`Jira:CredentialTarget`)

Never commit tokens/passwords/cookies/bearer values or full auth URLs.
Authorization headers are never logged; `SecretRedactor` covers Basic/Bearer
and common token JSON keys.

Remote URLs must be **HTTPS**. Non-HTTPS is allowed only when
`Jira:AllowLocalHttpTestServer=true` **and** the host is loopback (tests /
explicit local mock servers).

Permissions used: **read-only** issue/search GETs. No Jira writeback.

## Adapter

```text
IJiraTestCaseSource
  └─ JiraRestTestCaseSource  (mockable named HttpClient "jira")
       ├─ canonical key validation before network
       ├─ configurable field/plugin mapping
       ├─ cancellation + bounded timeouts
       ├─ pagination (search)
       ├─ retry only idempotent GETs (bounded)
       ├─ ETag / Last-Modified conditional revalidation
       ├─ response-size limits
       └─ rate-limit reporting (X-RateLimit-* / Retry-After)
```

## Raw cache

```text
automation/jira-cache/<jiraKey>/<revision-or-hash>/
  test-case.json          # normalized Phase 4D input only
  source-metadata.json    # redacted fetch metadata (etag, sha, redacted paths)
```

Writes are atomic (temp sibling + rename). Unused Jira fields are not persisted.

## CLI

```cmd
dotnet run --project src/DesktopAutomationAgent -- sync-key --key DEMO-1 --json
dotnet run --project src/DesktopAutomationAgent -- compile-key --key DEMO-1 --from-jira --json
dotnet run --project src/DesktopAutomationAgent -- run-key --key DEMO-1 --dry-run --json
dotnet run --project src/DesktopAutomationAgent -- run-suite --file automation/suites/phase4e-demo.json --dry-run --json
dotnet run --project src/DesktopAutomationAgent -- run-suite --file automation/suites/phase4e-demo.json --continue-on-failure --max-failures 2 --dry-run
```

`compile-key` without `--from-jira` remains offline (Phase 4D behavior).

## Suite execution

1. Validate suite + canonical keys
2. Never run disabled keys / disabled suites
3. Sequential per enabled key: compile (or reuse cache) → Phase 2 run
4. Compile failure skips execution for that key
5. Fail-fast default; optional `--continue-on-failure` with `--max-failures N`
6. Per-key Phase 2 run report under `automation/runs/<runId>/`
7. Atomic suite summary at `automation/runs/<suiteRunId>/suite-run.json`
   (key, compilation fingerprint, plan hash, runId, result, duration)

Suite run directories do **not** overwrite Phase 2 plan-run directories.

## Out of scope

Scheduling, notifications, parallel execution, AI, Jira writeback.

## Related

- [Phase 4D deterministic BDD compiler](phase4d-deterministic-bdd-compiler.md)
- [Phase 2 deterministic runner](phase2-deterministic-runner.md)