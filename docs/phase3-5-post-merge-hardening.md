# Phase 3.5 — Post-Merge Hardening and Contract Alignment

## Objective

Harden and align the merged Phase 0–3 foundation before Jira integration, BDD
compilation, Debug Mode, or AI work. This phase is compatibility and
operational readiness only — no new automation capabilities.

## Base SHA

`cadfa0597d0694719bc84a38e00d536230806d94` (`main` at branch creation; tag `v1.0.111`).

## Jira contract decision

**Assistive Mode is authoritative.**

Canonical pattern:

```text
^[A-Z][A-Z0-9_]{0,31}-[1-9][0-9]{0,15}$
```

Agent implementation: `DesktopAutomationAgent.Configuration.JiraKeyContract`
(compiled, culture-invariant, bounded regex timeout).

Validation order:

1. Canonical Jira contract
2. Optional `Suites:JiraKeyPattern` project-specific additional restriction

Configured patterns cannot broaden acceptance: canonical always runs first.
Suite files still require uppercase keys. Interactive Assistive input may trim
and uppercase in the existing driver implementation (unchanged).

Aligned locations: `SuiteOptions`, agent `appsettings.json`,
`automation/schemas/suite.schema.json`, workspace templates,
Phase 1 docs, and drift tests against `bdd-action-map.schema.json`.

## outputPath clarification

`StartRecordingRequest.OutputPath` / README now document **directory only**.
Runtime behavior is unchanged: `RecordingService.ResolveOutputDirectory()`
treats the value as a directory and writes `recording_<timestamp>.json` inside it.
No filename heuristics were added.

## Release / version strategy

- Keep automatic `v1.0.<github.run_number>` tags for compatibility.
- Calculate version once in `release.yml`:
  - numeric: `1.0.<run_number>`
  - tag: `v1.0.<run_number>`
- Inject into `dotnet publish` via `-p:Version` / `-p:InformationalVersion`.
- Project files use unmistakable local default `0.0.0-local` (overridden on release).
- PR validation publishes both packages as CI artifacts and **does not** create a GitHub release.

## Driver / agent release artifacts

| Artifact | Contents |
|---|---|
| `DesktopAutomationDriver.exe` | Windows x64 self-contained driver |
| `DesktopAutomationAgent-win-x64.zip` | Agent exe + `appsettings.json` + required runtime files |
| `SHA256SUMS.txt` | SHA-256 checksums |

Driver and agent remain separate processes; the agent talks to the driver over loopback HTTP.

## Shared-host security limitation

Documented in [`shared-host-driver-discovery-security.md`](shared-host-driver-discovery-security.md).
Trusted single-user workstations can continue. Secure bootstrap is a release
blocker before untrusted shared-host deployment. Not implemented in Phase 3.5.

## Manual Windows acceptance status

Checklist: [`assistive-bdd-recording.md`](assistive-bdd-recording.md) (13 original steps + immediate start/stop).
Results template: [`assistive-windows-acceptance-results.md`](assistive-windows-acceptance-results.md).

**Status for this PR environment:** `NOT RUN — interactive Windows validation required`.

## Explicit out-of-scope list

- Jira REST / authentication
- Gherkin/BDD parsing or BDD-to-plan compilation
- Suite execution / scheduling
- AI / Copilot / LLM / self-healing
- Debug Mode / Ctrl+D implementation
- Full UIA object-tree traversal
- Database connections
- Notifications / HTML reports
- Automatic promotion of candidate Page Objects
- Driver route / operation / locator-priority / playback / Assistive event-format changes
- Broad refactor of `UiService.cs` or `RecordingOverlayWindow.cs`
- Secure bootstrap redesign (documented only)

## Commands used for verification

```text
dotnet test src/DesktopAutomationAgent.Tests/DesktopAutomationAgent.Tests.csproj --configuration Release
dotnet test src/DesktopAutomationDriver.Tests/DesktopAutomationDriver.Tests.csproj --configuration Release
dotnet publish src/DesktopAutomationDriver/DesktopAutomationDriver.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true
dotnet publish src/DesktopAutomationAgent/DesktopAutomationAgent.csproj --configuration Release --runtime win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=false
git diff --check
```

Windows GitHub Actions is authoritative for driver tests when the local host is not Windows.

## Known remaining work

- Record interactive Assistive Windows acceptance results on a real desktop.
- Phase 4: deterministic BDD compiler (not started).
- Future secure bootstrap for untrusted shared hosts (design only in this phase).
- Future Debug Mode services (boundaries only; see `debug-mode-design-boundaries.md`).
