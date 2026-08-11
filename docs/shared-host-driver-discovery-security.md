# Shared-host driver discovery security boundary

## Current facts

- The driver binds its **main API to loopback by default**.
- Authenticated endpoints require the **Bearer token**.
- The **`GET /verify` bootstrap endpoint returns the token without authentication**.
- Loopback is **machine-local** but **not necessarily Windows-user isolated**. Another
  local process (including another interactive user session on some Citrix/RDS
  topologies) may be able to call loopback HTTP.
- Agent **username matching** prevents accidental connection to another user’s driver
  instance when multiple drivers are present.
- Username matching **does not protect the bearer token** from another local process
  that can reach `/verify` on loopback.
- This matters on **shared Citrix / RDS / VDI hosts**.

## What is acceptable today

Trusted **single-user workstation** usage can continue with the current `/verify`
bootstrap model.

## Release blocker

**Secure bootstrap is a release blocker before untrusted shared-host deployment.**

Do not treat username matching or loopback binding alone as multi-user token isolation.

## Recommended future secure-bootstrap design (not implemented here)

Phase 3.5 documents the direction only. Implementing it changes the discovery
contract and needs an isolated review.

1. Prefer a **per-user ACL-protected named pipe** or **per-user ACL-protected
   connection descriptor** for bootstrap, rather than unauthenticated HTTP token return.
2. Use the Windows **SID** (not a normalized username string) as the identity for ACL checks.
3. **Never return the bearer token from unauthenticated HTTP** when secure mode is enabled.
4. Ensure **atomic creation and cleanup** of the bootstrap object.
5. **No token logging** in secure mode (stdout banners, traces, CI logs).
6. Provide a **migration/compatibility switch** for existing `/verify` clients.
7. Add tests using **multiple Windows identities** (or equivalent security-boundary tests).

## Documentation / test hygiene

Do not expose any real token in tests, logs, documentation, or CI output.
Use placeholders such as `REPLACE_WITH_TOKEN_FROM_VERIFY` in examples.
