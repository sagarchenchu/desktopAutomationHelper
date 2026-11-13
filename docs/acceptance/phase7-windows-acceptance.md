# Phase 7 Windows acceptance

- Tip tested: `19e826905078514aae9e75fc197b15bec990a62f` (`19e8269`)
- Requested review tip: `62d4c4a89ac894c0f36a8a3a544cd1761a6fd2d0` (S1 blocked without typed proof emission)
- Enabling commit on tip: emits `executionState=notStarted` / `failureCode=locatorNotFoundBeforeInvoke` on locator/window misses
- Branch: `feature/agent-phase-7-bounded-recovery`
- PR: #147 (draft - not merged)
- When: 2026-11-13T16:09:33.9140411+00:00
- Overall: **PASS**
- Approved pages unchanged: `True`
- Recovery outside budget: `False`

| # | Scenario | Expected | Result | Plan hash | Evidence |
|---|---|---|---|---|---|
| 1 | delayed dialog | recover | PASS `passed` | `38a7e06b56509ddb10260f29bf932d608d82f9e5b5f1ebcb4b21234a790e0da6` | `phase7-delayed-dialog-result.json` |
| 2 | approved to alternate | recover | PASS `passed` | `c7fe31082f5175cc56b28767bc8c1c4025eeff5bdbd23817f87bbcd04cb58981` | `phase7-approved-alternate-result.json` |
| 3 | ambiguous | STOP | PASS `failed` | `a29b3f87d8ab57bfc1fa89d2efcc7ec2cf47bd60d748221bf8d8bf1693dd76ca` | `phase7-ambiguous-result.json` |
| 4 | unexpected window | STOP | PASS `passed` | `dcadaf3751a966e970f9555300f76cf6c03baca03d5d20a48325ea67c10bf8f1` | `phase7-unexpected-window-result.json` |
| 5 | assertion | STOP | PASS `failed` | `270bc15f209f29be122dfb193813d50b007e4efdaa2faa8a840fc67e5c7d5412` | `phase7-assertion-result.json` |
| 6 | only permitted + immutability | meta | PASS | n/a | `phase7-only-permitted-result.json` |