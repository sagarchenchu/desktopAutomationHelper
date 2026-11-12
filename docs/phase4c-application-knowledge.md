# Phase 4C — Versioned application knowledge

Convert multiple Phase 4B exploration runs into a deterministic, reviewable
application state graph while preserving separation between **candidates** and
**approved** knowledge (and never auto-promoting Page Objects).

## Directory model

```
automation/knowledge/
  candidates/<applicationId>/<buildFingerprint>/<knowledgeRunId>/
    application-knowledge.json
    state-graph.json
    workflows.json
    coverage.json
  approved/<applicationId>/<buildFingerprint>/
    application-knowledge.json
    state-graph.json
    workflows.json
    coverage.json
    knowledge-approval.json
```

Candidate contents are gitignored. Approved knowledge is tracked only when
sanitized and intentionally reviewed.

## Schemas

| Schema | Role |
|---|---|
| `application-knowledge.schema.json` | Envelope: fingerprint, sources, graph, coverage, approval |
| `application-state-graph.schema.json` | Windows, states, elements, transitions, dialogs |
| `application-workflow.schema.json` | Ordered successful transition sequences |
| `knowledge-approval.schema.json` | Human review record (`promotesPageObjects: false`) |

```mermaid
flowchart LR
  E1[Exploration run A] --> B[build-knowledge]
  E2[Exploration run B] --> B
  B --> C[candidates/...]
  C --> V[validate-knowledge]
  V --> A[approve-knowledge]
  A --> P[approved/...]
  C -.->|optional suggestions only| POC[object-repository/candidates]
  A -.->|never| POP[object-repository/pages]
```

## Fingerprinting

A knowledge version is identified by `applicationId` + `buildFingerprint`.

Exploration runs now seal a `build-identity.json` at crawl time. Knowledge merge
**fails closed** when that file is missing or incomplete, and never re-reads a
mutable policy file for identity (live policy files must still match the sealed
`policySha256` when present).

Fingerprint fields:

| Field | Notes |
|---|---|
| `applicationId` | Stable app slug |
| `executablePath` | Prefer live process path sealed at exploration |
| `productVersion` / `fileVersion` | Captured from live `FileVersionInfo` when available |
| `executableSha256` | Prefer live executable SHA-256 sealed at exploration |
| `configIdentity` | Relevant config/env **without secrets**; always includes `processName` |
| `driverVersion` | Required (from sealed identity) |
| `snapshotSchemaVersion` | Semantic snapshot contract |
| `explorerVersion` | Exploration artifact contract (`phase4b-1`) |
| `explorationPolicyHash` | Policy SHA-256 |

At least one **immutable build id** is required: `executableSha256`,
`productVersion`, or `fileVersion`. Same policy hash with a different executable
hash yields incompatible fingerprints (new `buildFingerprint` path).

**Incompatible fingerprints are never silently merged.** A fingerprint change
starts a new knowledge version under a new `buildFingerprint` path. Duplicate
exploration run ids / artifact hashes are rejected (no inflated observation counts).

## Stable IDs

Deterministic IDs are hashed from normalized semantic content:

- Exclude runtime refs, HWNDs, coordinates, timestamps, ordinal/depth suffixes
- Element IDs prefer `AutomationId` + `ControlType` (+ `ClassName`) over Name-only
- Collisions are detected and reported in `idCollisions`
- Content hashes for approval/compare exclude volatile first/last observed timestamps

## Merge rules

1. Validate exploration artifacts before merge
2. Sort runs by `explorationId` (deterministic; order-independent input)
3. Identical normalized transitions increase `observationCount` + evidence
4. Conflicting transitions (same source+action, different target/outcome) remain
   explicit distinct edges sharing a `conflictGroupId`
5. Missing controls in a later run do **not** auto-delete prior knowledge
6. Build-fingerprint mismatch rejects the merge

## Candidate / approved separation

| Action | Candidate | Approved | Page Objects |
|---|---|---|---|
| `build-knowledge` | writes under `knowledge/candidates` only | no | optional suggestions under `candidates/` only |
| `approve-knowledge` | reads candidates only | writes | **never** promotes |
| Active `pages/` / `repository.json` | untouched | untouched | untouched |

`--output` and approval `--source` must resolve under `knowledge/candidates`.
Paths under `knowledge/approved` or `object-repository/pages` are rejected.
Candidate builds emit `content-hash.txt` + `approval-template.json`. The
canonical content hash binds `knowledgeId` (excludes status/approval/timestamps).
Approved validate requires `content-hash.txt`, recomputes the hash, requires
`approval.sourceKnowledgeId == knowledgeId`, and cross-checks sibling
`knowledge-approval.json`, `state-graph.json`, `workflows.json`, and
`coverage.json` against the document. Absolute machine paths are stripped from
approved artifacts.

## CLI (offline — no driver / network)

```cmd
dotnet run --project src/DesktopAutomationAgent -- build-knowledge --application-id demo-app --exploration-run <id> --exploration-run <id2> --json
dotnet run --project src/DesktopAutomationAgent -- validate-knowledge --path <knowledge-dir> --json
dotnet run --project src/DesktopAutomationAgent -- compare-knowledge --left <a> --right <b> --json
dotnet run --project src/DesktopAutomationAgent -- search-knowledge --path <knowledge-dir> --query open --json
dotnet run --project src/DesktopAutomationAgent -- approve-knowledge --source <candidate-dir> --reviewer alice --reviewed-input approval.json --json
```

`search-knowledge` is **deterministic lexical search only** (no embeddings / vector DB).

`approve-knowledge` requires:

- `--reviewed-input` JSON with `decision=approve`
- matching `sourceKnowledgeHash` (content hash excluding volatile timestamps)
- matching `--reviewer`
- `promotesPageObjects: false` (always)

## Coverage meanings

| Field | Meaning |
|---|---|
| `windows` | Distinct semantic windows |
| `uniqueStates` | Distinct state fingerprints |
| `actionableElements` | Elements with locator evidence |
| `permittedTransitionsAttempted` | Sum of transition observation counts |
| `successfulTransitions` | success / newWindow / closedWindow / modalDialog observations |
| `noOpTransitions` | no-op observations |
| `failedTransitions` | failure / scope / timeout / stale observations |
| `deniedSamples` | Denied action samples from exploration summaries |
| `unresolved` | Unresolved + ambiguous elements |
| `conflicts` | Transition edges in a conflict group |
| `truncationOrBudgetLimits` | Stop reasons / budget warnings |

## Object repository protection

- Suggestions go only to `object-repository/candidates/`
- Never writes `object-repository/pages/`
- Never modifies `repository.json`
- Approved object refs may be linked from knowledge only after separate human validation

## Workflow continuity

Successful transitions are chained only while `previous.target == next.source`.
Discontinuities split into separate workflows. Validator enforces entry/terminal
alignment, contiguous ordinals, and source→target continuity.

## Limitations

- No AI / embeddings / vision
- No automatic Page Object promotion
- No silent rebuild across incompatible builds
- WinForms window class names are normalized by stripping runtime `.app.*` suffixes
- Product/file version enrichment depends on sealed exploration identity
- Workflow extraction covers continuous successful navigation segments only
- Knowledge merge does not re-query the live UI
- Exploration runs without `build-identity.json` cannot be merged
