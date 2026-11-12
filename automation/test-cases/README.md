# Offline Jira test-case exports (Phase 4D)

Place sanitized Jira-exported BDD JSON here (no credentials or raw production data).

- Schema: `../schemas/jira-test-case.schema.json`
- Example: `examples/DEMO-1.json`

Compile with:

```cmd
dotnet run --project src/DesktopAutomationAgent -- compile-key --key DEMO-1 --json
```
