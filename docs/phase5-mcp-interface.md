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
- File arguments must resolve **inside the workspace root** (path traversal rejected)
- Structured errors are redacted (`SecretRedactor`)
- Observation responses are **bounded** (depth/children/nodes + response character budget)
- Opaque observe refs are in-process only and expire (default 15 minutes)
- Tool descriptions distinguish **observation / navigation / mutation / execution / validation**
- There is **no** unrestricted `desktop_execute` tool

## Tools

| Tool | Category | Notes |
|---|---|---|
| `desktop_doctor` | observation | Readiness / discovery checks |
| `desktop_observe` | observation | Bounded semantic snapshot + opaque refs |
| `desktop_explore_dry_run` | observation | Policy candidates; no UI mutation |
| `desktop_explore` | navigation | Policy-bounded explorer |
| `desktop_search_knowledge` | observation | Lexical knowledge search |
| `desktop_validate_plan` | validation | Offline plan validation |
| `desktop_compile_key` | validation | BDD compile (`fromJira` optional sync) |
| `desktop_run_plan` | execution | Phase 2 runner (`dryRun` skips POST /ui) |
| `desktop_run_key` | execution | Compile + run key |
| `desktop_run_suite` | execution | Sequential suite orchestration |
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
