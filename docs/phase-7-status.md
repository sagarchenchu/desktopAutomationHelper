# Phase 7 status

## Required final report

Phase 7 status: **NOT READY** — five review findings at `bace213` addressed on a new tip; PR #147 remains draft. **Do not merge.** Windows acceptance evidence for the new tip is **pending** (do not claim 6/6 PASS).

Branch / base / head / PR:
- Branch: `feature/agent-phase-7-bounded-recovery`
- Base SHA: `cc93d333f95ba18b00c667121cba4dcb16822bba` (`origin/main` after Phase 6 #146 merge)
- Head: `41aadda833e2158875dd091562f6a60d576c7b66` / `41aadda` (fix commit `b67308a`)
- PR: https://github.com/sagarchenchu/desktopAutomationHelper/pull/147 — title `Phase 7: add bounded execution recovery`; **draft** (do not merge)

### Review tip `bace213` — five findings fixed

| # | Severity | Finding | Fix summary |
|---|---|---|---|
| 1 | P1 | Free-text non-execution proof | Typed `executionState: notStarted` / `failureCode: locatorNotFoundBeforeInvoke` only; free-text never authorizes mutation replay |
| 2 | P1 | Artifact write failures skip cleanup | `IOException` / `RecoveryPathException` (etc.) → recovery failure inside `TryRecoverAsync` / runner catch so deterministic cleanup still runs |
| 3 | P1 | Fingerprint used driver version | Live target process executable SHA via `IProcessIdentityProbe` / observe — never `driverStatus.Build` / `catalog.DriverVersion` |
| 4 | P2 | Deadline → Cancelled step missed BudgetExhausted | Ladder maps linked-token `Cancelled`/`ExecutionTimeout` step results to `BudgetExhausted`; test mirrors DriverUiClient conversion |
| 5 | P2 | Plan immutability fail-open | Missing / unreadable / null baseline plan ⇒ `PlanMutated=true` (fail-closed) |

Recoverable classifications:
- state not yet stable / bounded timing drift (**typed** non-execution proof for mutations; free-text timing phrases only for read/query)
- expected child window not yet active
- approved primary locator missing but approved alternate resolves uniquely
- known application-knowledge transition changed to another already-approved equivalent path
- explicitly approved candidate locator proposal

Non-recoverable classifications:
- assertion / business-result failure
- authentication / authorization failure
- destructive-action denial
- ambiguous locator (including multi-binding steps)
- application build mismatch without approved knowledge
- unexpected external process/window
- unknown operation
- compile / schema failure
- timed-out / uncertain mutation or protocol response without **typed** non-execution proof
- mutation failures with only free-text errors (e.g. “not found”, “timing”)
- recovery artifact / path failures (fail + cleanup)
- repeated failure after budget

Recovery ladder / budgets:
1. Fresh observation + bounded wait
2. Re-resolve same approved locator/object
3. Approved alternate locator (unique `finduia`)
4. Approved equivalent transition/path (never upgrades read→mutation when `AllowMutationReplay=false`)
5. Candidate AI repair proposal → stop for approval
6. Fail + original cleanup
- Budgets: `AllowMutationReplay=false`, `MaxTotalRecoveriesPerRun=12`, `MaxRecoveryAttemptsPerStep=8`, `MaxAdditionalDurationSeconds=60` (active recovery intervals only + linked CTS), `BoundedWaitMilliseconds=750`, `MaxCandidateArtifactBytes=65536`, `Provider=fake`
- Per-unit consume (probe/retry/model); no recursion; no recovery during cleanup
- Linked-token cancel that becomes a `Cancelled` step still surfaces `BudgetExhausted`

Plan immutability:
- Original plan bytes/SHA-256 recorded; recovery never rewrites the plan file
- Missing / unreadable plan after run start ⇒ `PlanMutated=true` (fail-closed)
- Repository before/after hashing unchanged by recovery

Candidate protection:
- Writes only under `object-repository/candidates/recovery/` with link-aware containment before/after create
- Never edits `pages/`, approved knowledge, or step registry; `autoPromoted=false`; redacted + size-bounded; fake provider in CI
- Oversized / duplicate / unsafe-junction writes become recovery failures (cleanup still runs)

Agent tests:
- `Phase7BoundedRecoveryTests` exercised locally after the five fixes (includes typed-proof, artifact-failure, cancelled-step deadline, plan fail-closed)
- Full Agent suite expected green on CI **Build and Test (Windows)** for the new tip

Driver tests:
- Covered by CI **Build and Test (Windows)** on the new tip

Windows scenarios:
- **PENDING** Windows acceptance re-run on the new tip after this push
- Prior store evidence (if any) applied to older tips and is **not** PR-visible proof for this tip
- `docs/acceptance/phase7-*` will be committed into the PR only after a successful re-run on the new tip — **do not claim 6/6 PASS** until then

Publish results:
- CI status: see latest push checks on PR #147 / branch tip

git diff --check:
- Clean on the tip that lands this fix

Known limitations:
- Equivalent transitions / fingerprints / approved candidates activate only when manifests/markers are present
- Candidate tier always stops for approval (never auto-executes model locators)
- Mutations never replay without typed driver non-execution proof (`AllowMutationReplay=false`)
- Not unrestricted self-healing
- Windows acceptance evidence for this tip is not yet in the PR

Exit gate:
Only explicitly recoverable controlled failures may recover, within budget, while the original plan and approved repositories remain unchanged.

**Exit gate: NOT MET** until CI is green on the new tip **and** Windows acceptance evidence is PR-visible. **PR #147 stays draft — do not merge.**

## Docs

- Repo: `docs/phase7-bounded-recovery.md`, `docs/phase-7-status.md`
- Store status: `/cursor/stores/bc-0b0638d3-f933-46b9-ac33-b6fb6b745c96/docs/phase-7-status.md`
- Acceptance: pending Windows re-run → then `docs/acceptance/phase7-*` in PR
