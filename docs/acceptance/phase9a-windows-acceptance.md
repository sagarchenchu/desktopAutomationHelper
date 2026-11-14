# Phase 9A — Windows multi-identity evidence

- Tip tested (GitHub-resolvable): [`b27e5f2fc8a6bc00d1f7a2f57a5edde991d49f08`](https://github.com/sagarchenchu/desktopAutomationHelper/commit/b27e5f2fc8a6bc00d1f7a2f57a5edde991d49f08) (`b27e5f2`)
- tipSha (single, unchanged across all result files): `b27e5f2fc8a6bc00d1f7a2f57a5edde991d49f08`
- Branch: `feature/security-phase-9a-secure-bootstrap`
- PR: #151 (**draft** — keep draft; do not merge)
- Host: `DESKTOP-5JD30DR` as `DESKTOP-5JD30DR\ADMIN` (`S-1-5-21-2097163997-1254372408-68186133-1001`)
- When: 2026-11-14T09:47:08.6254239+00:00
- OVERALL: **PASS** (with honest multi-user gap — see below)

## Multi-identity capability (honest gap)

| Check | Result |
|---|---|
| Genuine second interactive Windows user available? | **No** — only enabled local user is `ADMIN` |
| Create secondary local user (`net user ... /add`) | **Access denied** (System error 5; medium integrity; `Administrators` is deny-only) |
| Strongest ACL denial evidence used | Distinct well-known **LocalService** SID `S-1-5-19` vs interactive user SID — real Windows DACL `UnauthorizedAccessException` |
| Same-user mocks alone? | **No** — denial identity is a different Windows SID enforced by the kernel ACL |

**Gap:** Interactive multi-user isolation (User A vs User B on a shared Citrix/RDS host) is **not** proven on this machine. Do not treat LocalService-vs-interactive as a substitute for a second interactive account in security review.

## Scenarios

| # | Scenario | Expected | Result | Evidence |
|---|---|---|---|---|
| 1 | Intended SID succeeds (pipe + descriptor) | Current SID reads payload/token | **PASS** | `phase9a-01-intended-sid-result.json` |
| 2 | Other identity denied (pipe + descriptor) | LocalService-only ACL denies interactive SID | **PASS** | `phase9a-02-other-identity-acl-result.json` |
| 3 | Stale descriptor; concurrent starts; crash cleanup/restart | Concurrent refused; stale cleaned; restart OK | **PASS** | `phase9a-03-concurrent-crash-result.json` |
| 4 | Token never in secure logs/banners/errors | Secure banner omits bearer/operator tokens | **PASS** | `phase9a-04-token-absent-legacy-result.json` |
| 5 | Legacy only when explicitly configured + risk ack | `LegacyVerify` requires `AcknowledgeLegacyVerifyRisk` | **PASS** | `phase9a-04-token-absent-legacy-result.json`, VerifyController + agent tests |
| 6 | Agent auto-discovery uses secure bootstrap | Default `DiscoveryMode=Secure` → `secure-bootstrap` | **PASS** | `phase9a-05-agent-secure-discovery-result.json` |

## Test harness

- Live production-path runner: `SecureBootstrapHost` + `BootstrapAcl` (results in `phase9a-live-windows-evidence.json`)
- Filtered xUnit: `SecureBootstrapTests` + `VerifyControllerTests` — **15/15 PASS**
- Filtered xUnit: `DriverConnectionResolverTests` + `AgentReadinessServiceTests` — **22/22 PASS**
- Driver tests required `DOTNET_ROLL_FORWARD=LatestMajor` (host has Microsoft.AspNetCore.App 10.x, not 8.x)

## Summary

Secure SID ACL pipe/descriptor isolation, cleanup, token hygiene, legacy gating, and agent secure discovery are evidenced on tip `b27e5f2`. Interactive second-user denial remains a documented host gap for security review.
