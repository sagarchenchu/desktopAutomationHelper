# Phase 9A — Secure per-user driver discovery

## Objective

Close the shared-host discovery gap documented in
[`shared-host-driver-discovery-security.md`](shared-host-driver-discovery-security.md):
stop returning the bearer token from unauthenticated HTTP `/verify` by default,
and deliver connection material through a Windows SID ACL-protected channel.

## Default contract

| Component | Default |
|---|---|
| Driver `Bootstrap:Mode` | `Secure` |
| Agent `Driver:DiscoveryMode` | `Secure` |
| HTTP `/verify` token return | Disabled (403) |
| Token channel | Named pipe `DesktopAutomationDriver.bootstrap.{SID}` + `%LOCALAPPDATA%\DesktopAutomationDriver\bootstrap\connection.json` |
| Token logging | Disabled in secure mode |

## Legacy migration

`LegacyVerify` requires explicit mode **and** `AcknowledgeLegacyVerifyRisk=true`
on both driver and agent. Documented as unsafe on shared Citrix/RDS/VDI hosts.

## Tests

- Intended SID can read pipe/descriptor
- Other identity denied by Windows ACL (not same-user mocks alone)
- Stale descriptor detection + crash cleanup/restart
- Concurrent live starts refused
- Token absent from secure banners/errors
- Legacy only when explicitly configured
- Agent auto-discovery uses `secure-bootstrap`

Windows CI (`windows-latest`) is authoritative for ACL/pipe integration coverage.
Phase 9B (unattended desktop) is out of scope for this phase.
