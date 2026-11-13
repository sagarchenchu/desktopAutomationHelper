# Phase 7 — Bounded Execution Recovery

Phase 7 recovers from **limited, classified** runtime drift during deterministic
plan execution. It is **not** unrestricted self-healing: the original plan hash,
approved Page Objects, approved knowledge, and the step registry are never
silently rewritten.

```mermaid
flowchart TD
  F[Step failure] --> C{Classify}
  C -->|non-recoverable| X[Fail + original cleanup]
  C -->|recoverable + budget| L[Recovery ladder]
  L --> T1[1 Fresh observation + bounded wait]
  T1 --> T2[2 Re-resolve approved locator]
  T2 --> T3[3 Approved alternate locator]
  T3 --> T4[4 Approved equivalent transition]
  T4 --> T5[5 Candidate AI repair + stop for approval]
  T5 --> X
  T1 -->|success| P[Continue plan]
  T2 -->|success| P
  T3 -->|success| P
  T4 -->|success| P
```

## Recoverable classifications

| Classification | Meaning |
|---|---|
| `state-not-stable` | Timing drift / element not ready yet |
| `expected-child-window-not-active` | Expected dialog/child not yet active |
| `approved-alternate-locator` | Primary missing; approved alternate may resolve uniquely |
| `approved-equivalent-transition` | Approved knowledge maps to an equivalent path |
| `explicitly-approved-candidate-locator` | Previously reviewed candidate locator is available |

## Not automatically recoverable

| Classification | Meaning |
|---|---|
| Assertion / business-result failure | Never healed |
| Authentication / authorization | Never healed |
| Destructive-action denial | Never healed |
| Ambiguous locator | Stops; no guess |
| Application build mismatch without approved knowledge | Stops |
| Unexpected external process/window | Stops |
| Unknown operation | Stops |
| Compile / schema failure | Stops |
| Repeated failure after budget | Stops |

## Recovery ladder

1. Fresh observation evidence + **one** bounded wait when policy permits, then re-execute.
2. Re-resolve the same approved locator/object.
3. Try an existing **approved** alternate locator (`alternateLocators` on the Page Object) after unique `finduia` verification.
4. Use an existing **approved** equivalent knowledge transition/path.
5. Generate a **candidate** AI repair proposal under `object-repository/candidates/recovery/` and **stop for approval**.
6. Fail and execute the original `onFailureSteps` cleanup (unchanged, no recovery during cleanup).

## Budgets (`Recovery` settings)

| Setting | Default | Role |
|---|---|---|
| `Enabled` | `true` | Master switch (false = Phase 2 fail-fast) |
| `AllowMutationReplay` | `false` | When false, timed-out / uncertain mutation/invoke/input are never replayed |
| `MaxTotalRecoveriesPerRun` | `12` | Cap of recovery **work units** (each probe/retry/model call) across the run |
| `MaxRecoveryAttemptsPerStep` | `8` | Cap of work units per failing step |
| `MaxAdditionalDurationSeconds` | `60` | Accumulated **active recovery** wall-clock only (plan execution between units does not burn it; linked CTS cancels in-flight recovery work) |
| `BoundedWaitMilliseconds` | `750` | Tier-1 wait |
| `MaxCandidateArtifactBytes` | `65536` | Bound on redacted candidate repair JSON |
| `Provider` | `fake` | CI must use fake |

Additional hard rules:

- **No recursive recovery** (a recovery work unit cannot start another).
- **No recovery during cleanup** unless a future deterministic policy is added separately.
- Confidence metadata never authorizes execution or promotion.
- Multi-binding steps (`locator` + `parentLocator`/…) are **ambiguous** — recovery stops without rewriting.
- Approved repo SHA-256 is compared before/after; candidate writes use link-aware containment under `object-repository/candidates/` only.
- Optional manifests under `knowledge/approved/recovery/` wire equivalent transitions and application fingerprints; approved candidate markers live under `object-repository/candidates/recovery/*/approved.json`.

## Artifacts

Under `automation/runs/<runId>/recovery/`:

- Snapshots before/after recovery attempts
- Per-attempt records (classification, tier, outcome)
- `recovery-summary.json` with **original immutable plan hash**, tier used, candidate path, final result

The original plan file and its SHA-256 remain unchanged. A repaired future plan
requires a **separate reviewed** artifact/version.

## Candidate repairs and approval workflow

1. Fake (CI) or registered provider emits a candidate repair proposal.
2. Writer stores it only under `object-repository/candidates/recovery/<proposalId>/`.
3. Proposal includes evidence, confidence, old/new locators, application fingerprint,
   and a reproducible verification command.
4. **Never auto-promoted** into `object-repository/pages/`, approved knowledge, or the step registry.
5. Unique verification is required before human approval and promotion.

## Configuration

```json
"Recovery": {
  "Enabled": true,
  "AllowMutationReplay": false,
  "MaxTotalRecoveriesPerRun": 12,
  "MaxRecoveryAttemptsPerStep": 8,
  "MaxAdditionalDurationSeconds": 60,
  "BoundedWaitMilliseconds": 750,
  "MaxCandidateArtifactBytes": 65536,
  "Provider": "fake"
}
```

## Limitations

- Recovery cannot invent locators into the active repository.
- Ambiguous matches stop; they are never guessed.
- Assertion and destructive failures are never healed.
- Build/fingerprint mismatch without approved knowledge stops.
- Unexpected external windows/processes are out of scope for auto-recovery.
- This phase does **not** market or implement unrestricted self-healing.

## Exit gate

Only explicitly recoverable controlled failures may recover, within budget, while
the original plan and approved repositories remain unchanged.
