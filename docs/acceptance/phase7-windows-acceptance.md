# Phase 7 Windows acceptance

- Tip tested: `13c0f8af46e4e71f10194f3cae20801e970ca31c` (`13c0f8a`)
- Requested tip: `b37a1e9c207ba7fbfcc9b7e760d829aa80fcd6b7` (+ structured `found=false` on switchwindow window miss)
- Branch: `feature/agent-phase-7-bounded-recovery`
- PR: #147 (draft - not merged)
- When: 2026-11-13T17:00:34.6454003+00:00
- Overall: **PASS**
- Approved pages unchanged: `True`
- Recovery outside budget: `False`

## Scenario 4 (unexpected-window) - runner STOP (not a classifier probe)

- reportStatus: `failed` exitCode: `6` (plan failed; agent exit non-zero)
- runnerUnexpectedExternal: `True` — DeterministicPlanRunner PID/scope enforcement on `switchwindow` to foreign Notepad
- recoveryClassification: `UnexpectedExternalWindow` (recovery denied; `totalRecoveryAttempts: 0`)
- No separate classifier probe was used for PASS.

| # | Scenario | Expected | Result | Plan hash | Evidence |
|---|---|---|---|---|---|
| 1 | delayed dialog | recover | PASS `passed` | `38a7e06b56509ddb10260f29bf932d608d82f9e5b5f1ebcb4b21234a790e0da6` | `phase7-delayed-dialog-result.json` |
| 2 | approved to alternate | recover | PASS `passed` | `c7fe31082f5175cc56b28767bc8c1c4025eeff5bdbd23817f87bbcd04cb58981` | `phase7-approved-alternate-result.json` |
| 3 | ambiguous | STOP | PASS `failed` | `a29b3f87d8ab57bfc1fa89d2efcc7ec2cf47bd60d748221bf8d8bf1693dd76ca` | `phase7-ambiguous-result.json` |
| 4 | unexpected window | STOP via **runner** | PASS `failed` exit `6` | `f70b32edcc994dabaf21b2094ed1b3d0d22ddc76ff7b520365e85a9cd732ffcb` | `phase7-unexpected-window-result.json` |
| 5 | assertion | STOP | PASS `failed` | `270bc15f209f29be122dfb193813d50b007e4efdaa2faa8a840fc67e5c7d5412` | `phase7-assertion-result.json` |
| 6 | only permitted + immutability | meta | PASS | n/a | `phase7-only-permitted-result.json` |