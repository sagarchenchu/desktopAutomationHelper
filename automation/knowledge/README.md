# Application knowledge (Phase 4C)

Versioned application state graphs derived from Phase 4B exploration runs.

## Layout

```
knowledge/
  candidates/<applicationId>/<buildFingerprint>/<knowledgeRunId>/
  approved/<applicationId>/<buildFingerprint>/
```

- **candidates/** — gitignored merge outputs and intermediate artifacts
- **approved/** — human-reviewed knowledge only (sanitized examples may be tracked)

## Separation

- Candidate knowledge never silently becomes approved.
- `approve-knowledge` copies a reviewed candidate after verifying source hash + reviewer.
- Knowledge approval **never** promotes Page Objects (`promotesPageObjects` is always `false`).
- Page Object suggestions are written only under `object-repository/candidates/`.

## Offline CLI

```cmd
dotnet run --project src/DesktopAutomationAgent -- build-knowledge --application-id demo-app --exploration-run <id> --json
dotnet run --project src/DesktopAutomationAgent -- validate-knowledge --path <knowledge-dir> --json
dotnet run --project src/DesktopAutomationAgent -- compare-knowledge --left <a> --right <b> --json
dotnet run --project src/DesktopAutomationAgent -- search-knowledge --path <knowledge-dir> --query open --json
dotnet run --project src/DesktopAutomationAgent -- approve-knowledge --source <candidate-dir> --reviewer alice --reviewed-input <approval.json> --json
```

See `docs/phase4c-application-knowledge.md`.
