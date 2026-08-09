# Object repository (Phase 3)

The object repository stores **approved, versioned UI locators** for deterministic plan execution.
Phase 3 provides the core read/validate/resolve library only. Capture, verify, CLI, and plan
expansion are implemented in later phases.

## Layout

```text
object-repository/
  repository.json          # manifest (tracked)
  pages/                   # active page-object documents (tracked)
  candidates/              # draft page objects awaiting promotion (gitignored)
  captures/                # raw capture output from tooling (gitignored)
```

## Captures vs approved objects

| Area | Purpose | Git |
|------|---------|-----|
| `captures/` | Raw, machine-generated locator snapshots from a capture session | Ignored |
| `candidates/` | Human-reviewed drafts promoted from captures | Ignored |
| `pages/` | Active page-object JSON referenced by `repository.json` | Tracked |

**Captures are never executed directly.** They may include volatile fields (window handles,
coordinates, runtime IDs) and incomplete locators. Operators review captures, curate locators,
and **manually promote** approved definitions into `pages/` with `state: "active"`.

Promotion is intentional and explicit. The agent does not auto-promote captures or candidates.

## PII and secrets

- Do **not** store passwords, tokens, customer data, or other PII in page objects.
- Prefer stable `automationId` and structural locators over visible text that may contain names.
- Review capture output before promotion; redact or generalize sensitive `name` values.

## Locator rules (enforced by the agent)

Allowed locator fields:

- `automationId`
- `name` + `controlType`
- `className` + `controlType`
- `matchMode`: `exact`, `contains`, or `startswith`
- `foundIndex` (discouraged; adds fragility warnings)

Volatile fields (handles, coordinates, runtime IDs, bounding boxes, etc.) are **rejected**.

`automationId` alone is sufficient. Without it, `name` and `controlType` **or** `className` and
`controlType` are required. Name-only, className-only, or controlType-only locators fail validation.

## Manual promotion workflow

1. Run capture tooling (later phase) to write artifacts under `captures/`.
2. Review and edit locators; copy the page JSON into `candidates/` while iterating.
3. Set `state` to `active`, `source.kind` to `manual` or `approved`, and add the page to
   `repository.json` under `pages/`.
4. Commit only the manifest and `pages/` files. Never commit `captures/` or `candidates/`.

Active pages must not contain `source.kind: "capture"` elements.

## No AI in Phase 3

Phase 3 does **not** call AI providers, perform self-healing, or rewrite locators automatically.
All approved objects are human-reviewed. Later phases may add optional diagnostics, but execution
continues to use only validated, active repository entries.

## Identifiers

`repositoryId`, `pageId`, and element keys must match:

```text
^[a-z][a-z0-9-]{0,63}$
```

Object references use `pageId.elementId` (for example `login.submit-button`).

## Schemas

- `schemas/object-repository.schema.json` — manifest
- `schemas/page-object.schema.json` — page documents

Do not store secrets here.
