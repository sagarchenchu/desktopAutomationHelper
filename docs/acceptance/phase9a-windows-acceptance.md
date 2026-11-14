# Phase 9A — Windows multi-identity evidence (DNU-38)

- Tip tested (GitHub-resolvable): [`6938263d32fb045589353c69eecc033d79c597a2`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/6938263d32fb045589353c69eecc033d79c597a2) (`6938263`)
- tipSha (single, unchanged across all result files): `6938263d32fb045589353c69eecc033d79c597a2`
- Branch: `feature/security-phase-9a-secure-bootstrap`
- PR: #151 (**draft** — keep draft; do not merge)
- Host: `DESKTOP-5JD30DR` as `DESKTOP-5JD30DR\ADMIN` (`S-1-5-21-2097163997-1254372408-68186133-1001`)
- When: 2026-11-14T10:33:59.8030179+00:00
- CI: 4/4 green on tip `6938263` before evidence
- OVERALL: **PASS_WITH_GAP**
- productionMultiUserGate: **NOT_PROVEN**

## Multi-identity capability (honest gap)

| Check | Result |
|---|---|
| Genuine second interactive Windows user available? | **No** — only enabled local user is `ADMIN` |
| Create secondary local user (`net user ... /add`) | **Access denied** (System error 5; medium integrity; `Administrators` deny-only; Windows `sudo` disabled) |
| Production multi-user gate (User A publish → User B denied pipe+descriptor) | **NOT_PROVEN** |
| LocalService-only test ACL used as substitute? | **No** — explicitly not claimed as PASS for this gate |

**Gap:** Interactive multi-user isolation under **production** ACL (User A vs User B on a shared Citrix/RDS host) is **not** proven on this machine. Do not treat prior LocalService-vs-interactive ACL experiments as a substitute for a second interactive account in security review.

## Scenarios

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | Intended SID succeeds (pipe + descriptor) under production ACL | Current SID reads payload/token; pipe server PID matches | **PASS** | `phase9a-01-intended-sid-result.json` |
| 2 | Second interactive user denied under **production** ACL | User B denied pipe connect + descriptor read | **NOT_PROVEN** | `phase9a-02-other-identity-acl-result.json` |
| 3 | Concurrent reject; rejected dispose preserves descriptor; stale PID+creation-time cleanup | Concurrent refused; live descriptor preserved; restart OK | **PASS** | `phase9a-03-concurrent-crash-result.json` |
| 4 | Token never in secure logs/banners/errors | Secure banner omits bearer/operator tokens | **PASS** | `phase9a-04-token-absent-legacy-result.json` |
| 5 | Legacy only when explicitly configured + risk ack | `LegacyVerify` requires `AcknowledgeLegacyVerifyRisk` | **PASS** | `phase9a-04-token-absent-legacy-result.json` |
| 6 | Agent auto-discovery uses secure bootstrap | Default `DiscoveryMode=Secure` → `secure-bootstrap` | **PASS** | `phase9a-05-agent-secure-discovery-result.json` |

## Test harness

- Live production-ACL runner: `SecureBootstrapHost` + production `BootstrapAcl` (results in `phase9a-live-windows-evidence.json`)
- Filtered xUnit: `SecureBootstrapTests` + `VerifyControllerTests` — **15/15 PASS**
- Filtered xUnit: `DriverConnectionResolverTests` + `AgentReadinessServiceTests` — **22/22 PASS**
- Driver tests required `DOTNET_ROLL_FORWARD=LatestMajor` (host has Microsoft.AspNetCore.App 10.x, not 8.x)

## Summary

On tip `6938263` (CI green): intended SID production ACL path, pipe-server PID auth check, concurrent reject + rejected-dispose preserve, and stale PID+creation-time restart all PASS. The required second-interactive-user production ACL denial gate remains **NOT_PROVEN** due to host privilege limits.
