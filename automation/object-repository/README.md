# Object repository (Phase 3)

The object repository stores **approved, versioned UI locators** for deterministic plan execution.

## Layout

```text
object-repository/
  repository.json          # manifest (tracked)
  pages/                   # active page-object documents (tracked)
  candidates/              # draft page objects awaiting promotion (gitignored)
  captures/                # raw capture output from tooling (gitignored)
```

## CLI workflow

```bash
# Offline validation
dotnet run --project src/DesktopAutomationAgent -- validate-object-repository \
  --file automation/object-repository/repository.json

# Offline resolve one reference
dotnet run --project src/DesktopAutomationAgent -- resolve-object \
  --file automation/object-repository/repository.json --ref login.submit

# Live capture (dumpuia) — writes captures/ and candidates/
dotnet run --project src/DesktopAutomationAgent -- capture-page \
  --file automation/object-repository/repository.json \
  --page login --name "Login page" \
  [--view control|content|raw] \
  [--root activeWindow|processWindows|desktopChildren] \
  [--max-depth 8] [--max-children 200] [--include-offscreen] \
  [--json]

# Live verify (finduia) — all active objects, or filter
dotnet run --project src/DesktopAutomationAgent -- verify-object-repository \
  --file automation/object-repository/repository.json \
  [--page login | --ref login.submit] [--json]
```

Plans may reference repository objects via `$objectRef` in locator arguments when
`objectRepository` is set on the plan. See `docs/phase3-object-repository.md`.

## Captures vs approved objects

| Area | Purpose | Git |
|------|---------|-----|
| `captures/` | Raw, machine-generated locator snapshots from a capture session | Ignored |
| `candidates/` | Human-reviewed drafts promoted from captures | Ignored |
| `pages/` | Active page-object JSON referenced by `repository.json` | Tracked |

**Captures are never executed directly.** Operators review captures, curate locators,
and **manually promote** approved definitions into `pages/` with `state: "active"`.

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
`controlType` are required.

## Manual promotion workflow

1. Run `capture-page` to write artifacts under `captures/` and `candidates/`.
2. Review and edit locators; iterate in `candidates/` if needed.
3. Set `state` to `active`, `source.kind` to `manual` or `approved`, and add the page to
   `repository.json` under `pages/`.
4. Run `verify-object-repository` against the promoted page.
5. Commit only the manifest and `pages/` files. Never commit `captures/` or `candidates/`.

Active pages must not contain `source.kind: "capture"` elements.

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
