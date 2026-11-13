# Phase 5 — Desktop Automation MCP interface

Phase 5 exposes already-tested `DesktopAutomationAgent` application services through a
**local stdio MCP sidecar**. Handlers are a protocol adapter only: they do not
re-implement driver clients, validators, the BDD compiler, exploration policy, or
plan runners.

## Deployment

```text
MCP client (same user)
   │ stdio JSON-RPC
   ▼
DesktopAutomationAgent mcp-server
   │ existing DI services
   ▼
DesktopAutomationDriver (internal discovery) / workspace artifacts
```

Start:

```bash
dotnet run --project src/DesktopAutomationAgent -- mcp-server
```

- Prefer **stdio** (this is the only transport shipped in Phase 5)
- Local sidecar / same-user process tree
- **No** unauthenticated network listener
- Driver bearer tokens / Authorization headers are **never** returned via MCP
- Logs go to **stderr**; stdout is reserved for JSON-RPC

## Security assumptions

- The MCP client and agent run as the **same local user**
- Driver discovery remains the existing internal/local mechanism
- File arguments must resolve **inside the workspace root** after link-aware expansion (symlinks / junctions / chained reparse points rejected when they escape)
- Structured errors are redacted (`SecretRedactor`); JSON string values stay valid after redaction
- Observation responses are **bounded** (depth/children/nodes + progressive response character budget; always valid JSON)
- Opaque observe refs are in-process only and expire (default 15 minutes)
- Tool descriptions distinguish **observation / navigation / mutation / execution / validation**
- There is **no** unrestricted `desktop_execute` tool

## Tools

| Tool | Category | Notes |
|---|---|---|
| `desktop_doctor` | observation | Readiness / discovery checks |
| `desktop_observe` | observation | Bounded semantic snapshot + opaque refs |
| `desktop_explore_dry_run` | observation | Policy candidates; no UI mutation (not ReadOnly — may write artifact stubs) |
| `desktop_explore` | navigation | Policy-bounded explorer |
| `desktop_search_knowledge` | observation | Lexical knowledge search |
| `desktop_validate_plan` | validation | Offline plan validation |
| `desktop_compile_key` | validation | BDD compile (not ReadOnly — writes plan cache); `fromJira` is OpenWorld |
| `desktop_run_plan` | execution | Phase 2 runner; `dryRun` defaults to **true** (omission never live-runs) |
| `desktop_run_key` | execution | Compile + run key; OpenWorld; `dryRun` defaults to **true** |
| `desktop_run_suite` | execution | Sequential suite orchestration; OpenWorld; `dryRun` defaults to **true** |
| `desktop_get_run_summary` | observation | Compact sanitized run / suite-run summary |

### Errors

Structured tool errors use codes such as:

`validation`, `path_traversal`, `stale_ref`, `ambiguous`, `denied`,
`driver_unavailable`, `cancelled`, `truncated`, `internal`.

## Resources (optional)

- `desktop://catalog/summary` — sanitized step-registry summary
- `desktop://knowledge/{knowledgeId}/summary` — sanitized knowledge counts
- `desktop://runs/{runId}/summary` — compact run summary

No secrets and no unrestricted raw artifacts.

## Client config

See `automation/config/mcp-client.example.json` (placeholders only).

## Out of scope

Model integration, autonomous planning, self-healing, and public remote MCP.
