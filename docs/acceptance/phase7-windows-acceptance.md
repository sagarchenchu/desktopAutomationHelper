# Phase 7 Windows acceptance

- Branch: `feature/agent-phase-7-bounded-recovery`
- PR: #147 (draft — **do not merge**)
- Prior tip with partial evidence: `19e8269` / `70067f3` (scenario 4 was **not** a real runner STOP — superseded)
- Current code tip under review: see branch HEAD after ae41ca0 follow-up fixes

## Status

**NOT claiming 6/6 PASS.** Scenario 4 (unexpected external window) must be re-run as an end-to-end runner STOP:

1. Launch target app (pin `launchedProcessId`)
2. `switchwindow` to a foreign process (e.g. Notepad)
3. `DeterministicPlanRunner` itself fails that step with `Unexpected external window … outside approved launched process`
4. Classifier: `UnexpectedExternalWindow`, `mayAttemptRecovery=false`
5. Plan hash unchanged; cleanup still runs

Do **not** mark scenario 4 PASS via a separate classifier probe while the plan `status: passed`.

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | delayed dialog | recover | pending re-run on new tip | prior JSON at tip `19e8269` only |
| 2 | approved to alternate | recover | pending re-run on new tip | prior JSON at tip `19e8269` only |
| 3 | ambiguous | STOP | pending re-run on new tip | prior JSON at tip `19e8269` only |
| 4 | unexpected window | STOP via **runner** | **PENDING** real e2e | unit-covered; Windows evidence TBD |
| 5 | assertion | STOP | pending re-run on new tip | prior JSON at tip `19e8269` only |
| 6 | only permitted + immutability | meta | pending re-run on new tip | prior JSON at tip `19e8269` only |

## Enabling code (this tip)

- Typed proof only from locator-resolution path (`found=false` / `matchCount=0` / `UiResolutionException`) — never free-text
- Runner pins launch PID and fails `switchwindow` to foreign PID
- Fingerprint after launch / at recovery; skip when recovery disabled
- Zero-duration budget on ladder entry → `BudgetExhausted`
