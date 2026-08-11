# Phase 1 — Desktop Automation Agent Foundation

## Goal

Introduce an independent .NET 8 command-line agent that treats
`DesktopAutomationDriver` as an external HTTP service. Phase 1 establishes
configuration, safe driver discovery, workspace layout, suite-manifest
validation, and readiness checks. It does **not** execute UI operations.

## Agent / driver HTTP boundary

| Allowed | Forbidden |
|---------|-----------|
| `GET /verify` | `POST /ui` |
| `GET /status` | Any session/launch/playback route |
| `GET /ui/operations` | Project references to `DesktopAutomationDriver` |
| Agent-owned DTOs for those three endpoints | Copying driver `UiRequest` or internals |

The agent project targets `net8.0` (not `net8.0-windows`) and never references
FlaUI, WinForms, or driver assemblies.

## Configuration precedence

1. Checked-in `src/DesktopAutomationAgent/appsettings.json`
2. Optional local file `automation/config/agentsettings.local.json` (gitignored)
3. Environment variables with prefix `DA_AGENT__`
4. Command-line configuration arguments (`--Driver:BaseUrl=...`)

Examples:

```bash
export DA_AGENT__DRIVER__BASEURL=http://127.0.0.1:33201
export DA_AGENT__DRIVER__BEARERTOKEN='<token-from-verify>'
export DA_AGENT__WORKSPACE__ROOT=automation
```

Copy `automation/config/agentsettings.example.json` to
`agentsettings.local.json` for machine-local secrets. Never commit real tokens.

## Safe driver discovery

### Explicit connection

Supply **both** `Driver:BaseUrl` and `Driver:BearerToken`. Providing only one
is a configuration error (exit code `2`).

### Verify-endpoint discovery

When both explicit values are absent, the agent calls `Driver:VerifyUrl`
(default `http://localhost:9102/verify`), reads `username`, `port`, `token`,
and `authorizationHeader`, then builds:

`http://127.0.0.1:{port}/`

Verify discovery is **loopback-only**. A remote `VerifyUrl` is rejected even when
`AllowRemoteDriver=true`. For a remote driver, configure `BaseUrl` and
`BearerToken` explicitly. `AllowRemoteDriver` only relaxes explicit BaseUrl checks.

`doctor --json` writes a single JSON document to stdout. Structured logs always
go to stderr so CI parsers are not polluted.

### Citrix username validation

Because port 9102 may belong to another user on a shared Citrix/RDS host, the
discovered username is compared with `Environment.UserName`
(case-insensitive; domain/`UPN` suffixes stripped). On mismatch the agent:

- rejects the connection
- does not use the returned token
- asks the user to set BaseUrl + BearerToken explicitly

### Loopback enforcement

Remote hosts are rejected unless `Driver:AllowRemoteDriver=true`.

## Workspace layout

```text
automation/
  config/                 # examples + ignored local settings
  schemas/suite.schema.json
  suites/smoke.json
  suites/regression.json
  plans/                  # reserved (later phases)
  object-repository/      # reserved (later phases)
  runs/                   # generated evidence (gitignored contents)
```

`init` creates missing directories/templates and never overwrites existing files.
All resolved suite paths must remain inside the workspace root; `../` is rejected.

## Suite manifest format

```json
{
  "schemaVersion": 1,
  "name": "smoke",
  "enabled": true,
  "testCases": [
    { "jiraKey": "SAMPLE-1", "enabled": true }
  ]
}
```

Validation rules:

- `schemaVersion` must be `1`
- `name` required
- `testCases` required (may be empty)
- Jira keys must match the canonical Assistive contract `^[A-Z][A-Z0-9_]{0,31}-[1-9][0-9]{0,15}$` (suite files require uppercase). `Suites:JiraKeyPattern` is an optional project-specific additional restriction applied after the canonical check.
- duplicate keys fail
- disabled entries remain syntactically valid but are excluded from the effective selection
- errors identify file + `testCases[i]`
- no Jira network calls in Phase 1

## CLI commands and exit codes

| Command | Requires driver? | Purpose |
|---------|------------------|---------|
| `init` | No | Idempotent workspace bootstrap |
| `validate-suite --file <path>` | No | Validate one suite file |
| `validate-keys --keys A-1,B-2` | No | Validate ad-hoc keys |
| `doctor` | Yes (HTTP only) | Full readiness check |
| `doctor --json` | Yes | Machine-readable readiness for CI |

Exit codes:

| Code | Meaning |
|------|---------|
| 0 | Success |
| 2 | Usage or configuration error |
| 3 | Driver unavailable or unsafe discovery |
| 4 | Authentication or catalog incompatibility |
| 5 | Suite or workspace validation failure |

`doctor` calls only `GET /status` and `GET /ui/operations` after discovery. It
never launches apps, creates sessions, or calls `POST /ui`.

## Secret-handling rules

- Never log, print, or serialize bearer tokens or full `Authorization` headers
- Redact secrets from exceptions and readiness JSON
- Prefer loopback URLs
- Keep local secret files gitignored

## Phase 1 limitations

Not implemented here (later phases):

- Jira authentication / BDD fetch
- Gherkin parsing and command compilation
- AI providers / self-healing
- Executing driver UI operations
- Object-repository generation
- Database access
- Scheduling, notifications, evidence collection, parallel execution

## How later phases connect

| Later phase | Builds on |
|-------------|-----------|
| Jira + BDD | Suite manifests + key validation |
| Compiler | Workspace `plans/` + catalog from `/ui/operations` |
| Execution | Authenticated driver client patterns (still HTTP-only) |
| Object repository | `automation/object-repository/` |
| AI fallback | Failure evidence under `automation/runs/` |
| Reporting | Stable `doctor --json` exit codes and run artifacts |

## Build / test

```bash
dotnet restore DesktopAutomationHelper.slnx
dotnet build DesktopAutomationHelper.slnx --configuration Release
dotnet test src/DesktopAutomationAgent.Tests/DesktopAutomationAgent.Tests.csproj --configuration Release
dotnet run --project src/DesktopAutomationAgent -- init
dotnet run --project src/DesktopAutomationAgent -- validate-suite --file automation/suites/smoke.json
dotnet run --project src/DesktopAutomationAgent -- validate-keys --keys SAMPLE-1,SAMPLE-2
```
