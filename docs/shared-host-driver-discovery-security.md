# Shared-host driver discovery security boundary

## Current facts (Phase 9A)

- The driver binds its **main API to loopback by default**.
- Authenticated endpoints require the **Bearer token**.
- **Secure bootstrap is the default** (`Bootstrap:Mode=Secure` / `Driver:DiscoveryMode=Secure`).
- In secure mode the bearer token is delivered only via a **per-user Windows SID
  ACL-protected named pipe** and an **ACL-protected connection descriptor** under
  `%LOCALAPPDATA%\DesktopAutomationDriver\bootstrap\connection.json`.
- Unauthenticated **`GET /verify` never returns the bearer token in secure mode**
  (HTTP 403 `secure_bootstrap_required`).
- Loopback is **machine-local** but **not necessarily Windows-user isolated**. Another
  local process (including another interactive user session on some Citrix/RDS
  topologies) may be able to call loopback HTTP — which is why HTTP token return is
  disabled by default.
- Agent discovery matches the **Windows SID** on the secure channel (not username alone).
- Username matching remains a secondary guard for legacy `/verify` only.

## Secure bootstrap contract

| Channel | Identity | Token returned? |
|---|---|---|
| Named pipe `DesktopAutomationDriver.bootstrap.{SID}` | DACL: user SID + SYSTEM + Administrators | Yes (to allowed identities only) |
| Descriptor `%LOCALAPPDATA%\...\bootstrap\connection.json` | NTFS ACL: user SID + SYSTEM + Administrators | Yes (to allowed identities only) |
| `GET /verify` (main or probe 9102) | none | **No** in Secure mode |

Additional properties:

- **Atomic create/cleanup** of the descriptor: temp file is created with a restrictive
  `FileSecurity` **before** token bytes are written (crash must not leave a
  world-readable secret), then replaced into place; deleted only by the host that
  published it.
- **Exclusive per-user lock** (`host.lock`) serializes publication; a rejected second
  start must not delete the live descriptor on dispose.
- **Named-pipe server authentication**: the agent resolves the server process via
  `GetNamedPipeServerProcessId` and verifies the process token SID before trusting
  payload fields (SID / PID / port / token).
- **Stale detection** via owning **PID + process creation time** (PID reuse safe);
  crash leftovers are removed on next start.
- **Concurrent starts** refuse to overwrite a live descriptor.
- **No token logging** in secure mode (startup banner, traces, errors).

## Legacy `/verify` migration switch

Trusted **single-user workstation** clients that still call HTTP `/verify` can opt in:

**Driver `appsettings.json`:**

```json
"Bootstrap": {
  "Mode": "LegacyVerify",
  "AcknowledgeLegacyVerifyRisk": true
}
```

**Agent `appsettings.json`:**

```json
"Driver": {
  "DiscoveryMode": "LegacyVerify",
  "AcknowledgeLegacyVerifyRisk": true,
  "VerifyUrl": "http://localhost:9102/verify"
}
```

Both `Mode`/`DiscoveryMode` **and** `AcknowledgeLegacyVerifyRisk=true` are required.
Secure mode ignores the acknowledgement flag and never returns the token over HTTP.

### Legacy risk (shared-host)

**LegacyVerify is unsafe on shared Citrix / RDS / VDI hosts.** Any local process that
can reach loopback may call `/verify` and obtain the bearer token. Username matching
does **not** protect the token. Do not enable LegacyVerify for untrusted multi-user
hosts.

## Release posture

Secure bootstrap is required before untrusted shared-host deployment.
Do not treat username matching or loopback binding alone as multi-user token isolation.

## Documentation / test hygiene

Do not expose any real token in tests, logs, documentation, or CI output.
Use placeholders such as `REPLACE_WITH_TOKEN_FROM_VERIFY` (legacy) or
`REPLACE_WITH_TOKEN_FROM_SECURE_BOOTSTRAP` in examples.
