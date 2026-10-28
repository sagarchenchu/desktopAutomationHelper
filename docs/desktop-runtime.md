# Desktop execution runtime: first implementation

Status: **experimental diagnostics foundation (part of U0)**. There is no virtual
display driver, session broker, supervised `run` command, or proven execution
without an RDP viewer in this increment.

## Architecture decision

Build the runtime as an independent component, initially in this repository.
DesktopAutomationHelper is its first consumer. Keep its dependencies pointed
toward the runtime; the runtime must not reference the driver, agent, FlaUI,
pywinauto, an LLM, or test-specific assertions. Once the feasibility gate passes,
it can move to a separate repository with independent releases.

| Component | Responsibility |
| --- | --- |
| `DesktopExecutionRuntime` | Environment observations and conservative preflight policy; no third-party packages |
| `DesktopRuntime` | Language-neutral CLI; currently `doctor [--json]` only |
| Driver `RuntimeController` | Adapter that observes the driver's own Windows session through the shared library |
| Existing driver and agent | UI actions, test steps, assertions, and test reporting |
| Future session/display providers | Session worker lifecycle and existing VM graphics or a supported virtual display |

A virtual monitor supplies a display surface. It does not create, log in, unlock,
or make a Windows session interactive. The target app and UI worker must run in
the intended user session. A Jenkins service in Session 0 cannot solve that by
launching another ordinary child process. Session provisioning is a later,
explicitly tested capability.

## Use the diagnostic CLI

Build on a machine with the .NET 8 SDK:

```powershell
dotnet publish src/DesktopRuntime/DesktopRuntime.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o publish/runtime
& ./publish/runtime/DesktopRuntime.exe doctor --json > runtime.json
$doctorExit = $LASTEXITCODE
Get-Content runtime.json
```

The Build and Test workflow also produces the
`DesktopRuntime-experimental-win-x64` artifact. This is separate from the
driver and agent packages and is not yet included in tagged releases.

| Exit code | Meaning |
| --- | --- |
| `0` | Passive preflight checks passed; actual input/capture/test execution is still unverified |
| `2` | Preflight blocked, including unsupported operating systems or unknown observations |
| `64` | Invalid command or arguments |

The CLI can be called from Python, Java, PowerShell, Jenkins, or any tool that
can execute a process and read JSON. The Windows probe is platform-gated; Linux
returns `UNSUPPORTED_PLATFORM` without attempting Win32 calls.

Run it **inside the intended UI worker session**. Running it from a Jenkins
service diagnoses the service process, not a different logged-in user's desktop.
Collect both when troubleshooting a service/worker mismatch. Do not run as SYSTEM
or disable locking policies to make the report pass.

## Driver integration

```http
GET /runtime/diagnostics
Authorization: Bearer <existing-driver-token>
```

This uses the existing WebDriver envelope: the diagnostics are in `value`.
HTTP 200 and envelope `status: 0` mean the query succeeded. Inspect
`value.preflightPassed` and `value.blockingReasons` for the desktop result.
The legacy `/status` response is unchanged and remains a liveness signal.
The agent's existing `doctor` command is also unchanged; it does not automatically
call this endpoint in this increment.

The probe reads the process/session ID, WTS connection state, window-station
name/visibility, calling thread's desktop, input desktop, whether the thread's
desktop receives input, monitor count, and virtual-screen dimensions. Dimensions
are `GetSystemMetrics` values in the caller's DPI context, not a guarantee of
physical pixels. Reports include UTC capture time, OS description, and native
API error codes. Unknown observations block preflight.

`schemaVersion: "1.0"` identifies the report contract. Consumers should tolerate
additional JSON properties and treat unfamiliar blocking codes as failures.

| Blocking reason | Interpretation |
| --- | --- |
| `UNSUPPORTED_PLATFORM` | Current process is not on Windows |
| `INTERACTIVE_SESSION_UNAVAILABLE` | Session ID is missing or Session 0 |
| `SESSION_NOT_ACTIVE` | WTS state is unknown or not Active; disconnected sessions are conservatively blocked |
| `INTERACTIVE_WINDOW_STATION_UNAVAILABLE` | Station is not visible WinSta0, or its state is unknown |
| `INPUT_DESKTOP_UNAVAILABLE` | Input desktop is unavailable/different, or thread desktop does not observably receive input |
| `DISPLAY_UNAVAILABLE` | Monitor count or virtual-screen dimensions are missing/nonpositive |
| `PROBE_FAILED` | One or more native observations failed; inspect `snapshot.probeFailures` |

These observations are not atomic. A user can lock/disconnect the session after
the probe; the future worker must monitor changes during execution. An active
WTS session and an open desktop handle alone do not prove usability. Microsoft
documents that `OpenInputDesktop` can return a handle in a disconnected session.
No probe here calls `SwitchDesktop`, sends input, captures a screen, changes DPI,
or changes display/session configuration.

Every report explicitly leaves these capabilities unverified:

- `INPUT_INJECTION`: a deliberate fixture action and observed state change are required.
- `SCREEN_CAPTURE`: a meaningful fixture image must be checked, beyond file creation.
- `TARGET_SESSION_AND_INTEGRITY`: app/driver session and privilege compatibility are not inspected yet.
- `VIEWER_INDEPENDENT_EXECUTION`: requires the actual no-viewer acceptance experiment.

The endpoint inherits existing bearer authentication. Existing unauthenticated
`/verify` token discovery and startup token logging remain a documented
[shared-host limitation](shared-host-driver-discovery-security.md). This
increment does not authorize deployment to an untrusted shared host.

## Windows feasibility gate

Use an isolated test machine with a provisioned user session. For each row,
collect CLI output from that session, driver endpoint output, exact OS/build,
display provider, resolution/DPI, driver/app versions, timestamps, and test
results. Label the environment manually; the probe does not infer whether an
RDP viewer is visible or minimized.

| Environment | Evidence required |
| --- | --- |
| RDP connected and visible | Working baseline with real input and meaningful screenshot |
| RDP minimized | Same scenario; identify failures without assuming the desktop disappeared |
| RDP disconnected | Record actual behavior; preflight intentionally blocks disconnected WTS state |
| Provisioned console with no viewer connected | Required target: complete test scenario and useful screenshots |
| No physical monitor | Repeat console scenario using actual VM graphics or a chosen virtual display |
| Test session locked or logged off | Infrastructure failure, never a false passed test |
| Jenkins service versus user worker | Confirm different session evidence and correct worker placement |

The scenario must type into an enabled field, verify a disabled field, choose a
dropdown value, add a record, verify its grid row, update it, verify the changed
row, delete it, and verify absence. Include a deliberately incorrect assertion
to prove failure reporting. Run it with DesktopAutomationHelper and a small
pywinauto suite before claiming tool independence. Record repetitions and failures;
the proposed release gate is 30 successful nominal runs per supported profile.
Hosted CI compiling the code does not satisfy this gate.

For Jenkins, archive `runtime.json` and test artifacts in `post { always { ... } }`.
Keep infrastructure failure separate from test assertion failure. Do not add an
automatic retry around an uncertain add/update/delete operation.

## Delivery backlog

| Stage | Work and completion gate |
| --- | --- |
| U0, started here | Passive diagnostics. Remaining: opt-in agent doctor adapter, target session/integrity checks, input-result evidence |
| U1 | WinForms/WPF fixtures and the environment matrix above; prove execution without an open viewer |
| U2 | Attach-mode worker, exclusive session/desktop lease, timeout/cancellation, owned-process supervision and cleanup; generic CLI `run` contract |
| U3 | Session provisioning and display lifecycle; prefer existing VM graphics, evaluate a custom IDD only if needed; reboot recovery |
| U4 | Secure bootstrap, installer/versioned packages, Jenkins integration, combined runtime/test evidence |
| U5 | Supported Citrix/RDS configurations and concurrency across isolated sessions/VMs |
| U6 | Reproducible acceptance results, supported OS/display matrix, independent repository/release decision |

U-numbers are separate from the existing agent phase numbering. Extra monitors
do not isolate focus or input for parallel tests. Do not promise arbitrary locked
desktop automation or universal compatibility. Server 2019, newer Windows, and
Citrix must each be verified before selecting a display-driver implementation.

## Primary references

- [Interactive services and Session 0](https://learn.microsoft.com/en-us/windows/win32/services/interactive-services)
- [OpenInputDesktop, including disconnected-session behavior](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-openinputdesktop)
- [GetUserObjectInformationW: visibility, object names, and UOI_IO](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getuserobjectinformationw)
- [WTSQuerySessionInformationW](https://learn.microsoft.com/en-us/windows/win32/api/wtsapi32/nf-wtsapi32-wtsquerysessioninformationw)
- [Indirect display driver model](https://learn.microsoft.com/en-us/windows-hardware/drivers/display/indirect-display-driver-model-overview)
