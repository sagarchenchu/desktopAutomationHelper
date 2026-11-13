# Phase 6 — Guarded AI Semantic Planner and Retrieval

Phase 6 turns a natural-language testing goal into a **candidate Phase 2 plan**
through a guarded pipeline. The model never calls `DesktopAutomationDriver`.

```mermaid
flowchart LR
  G[NL goal] --> C[Canonicalize + injection guard]
  C --> R[Lexical retrieval]
  R --> M[IPlanningModelClient]
  M --> V[Schema + approved refs + risk]
  V --> P[Compile Phase 2 plan]
  P --> A[Artifacts + audit hashes]
  A --> X{Approval mode}
  X -->|proposal-only / approval-required / denied| Stop[No POST /ui]
  X -->|auto-run-known-safe + --execute| Run[Phase 2 runner]
```

## Boundaries

| Allowed | Forbidden |
|---|---|
| Retrieve small approved knowledge / PO / step / op subsets | Entire UIA trees, secrets, editable values, customer data |
| Typed JSON plan proposals | Model-selected unrestricted exploration |
| Deterministic validate / compile | Self-healing, vision/OCR, repo auto-update |
| Fake deterministic model in CI | Provider credentials in plans or UI/Jira text |
| Explicit `--execute` after auto-run-known-safe | Proposal-only execution; validator bypass |

`DesktopAutomationDriver` has **no** planning/provider dependency.

## Provider

- Interface: `IPlanningModelClient`
- Default / CI: `FakePlanningModelClient` (`Planning:Provider=fake`)
- Credentials only from env/secret config (`Planning:ApiKeyEnvironmentVariable`)
- Bounded timeout, token/output budgets, and cancellation

## Retrieval

Deterministic lexical search + filters first (embeddings optional, never source of truth):

- Approved knowledge under `knowledge/approved/...`
- Object repository pages
- Step registry mappings
- Operation catalog (offline file or minimal safe catalog)

Retrieval records IDs + content hashes in `retrieval-manifest.json`.

## Proposal schema

Checked-in: `automation/schemas/plan-proposal.schema.json`.

Required: goal, assumptions, unresolvedQuestions, ordered steps, object/knowledge refs,
expected state/postcondition, risk. Confidence is **metadata only** and never
authorizes execution. Malformed / extra / unknown / prohibited / ambiguous /
unsupported proposals are rejected.

## Approval modes

| Mode | Behavior |
|---|---|
| `proposal-only` (default) | Never executes / never `POST /ui` |
| `auto-run-known-safe` | Only when all mapped/approved/low-risk and validators pass; still needs `--execute` |
| `approval-required` | Data entry / transactional / uncertain — candidate may exist, no auto-run |
| `denied` | Destructive / prohibited / unresolved / injection / hallucinated / stale |

## Prompt injection

UI / Jira / app text is untrusted. System policy is not overridable. Requests for
credentials, shell, filesystem, network, or unrestricted driver access are rejected.
Embedded UI/Jira policy overrides fail closed before the model path proceeds.

## Artifacts

```
automation/planning/<planningId>/
  request-summary.json
  retrieval-manifest.json
  proposal.json
  validation.json
  candidate.plan.json
```

Redacted, bounded, non-overwriting (atomic temp + move). Full provider payloads are
not persisted unless a sanitized fixture.

## CLI

```cmd
dotnet run --project src/DesktopAutomationAgent -- propose-plan --goal "List visible windows" --json
dotnet run --project src/DesktopAutomationAgent -- propose-plan --goal "..." --knowledge knowledge/approved/<app>/<fp> --catalog planning/evals/catalog.json --json
```

## Eval set

Sanitized cases under `automation/planning/evals/`. CI uses the fake model.
Exit gate: valid goals produce candidates; unsafe / ambiguous / hallucinated /
stale / injected cases cannot reach the driver.

## Out of scope

Self-healing, model-selected unrestricted exploration, repo auto-update, vision/OCR.
