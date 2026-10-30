# DesktopRuntime single-monitor VDD prototype

This increment implements a native UMDF/IddCx driver, an administrator lifetime
utility, two GUI fixtures and an acceptance script. It is an **experimental
console-display prototype**. Its Windows build pipeline creates an **unsigned**
package; a successful build is not proof of installation or no-viewer automation.

The intended minimum is x64 Windows 10 1809 / Windows Server 2019 Desktop
Experience (17763), IddCx 1.3 and UMDF 2.25. Build headers are IddCx 1.4 with a
1.3 minimum. These are compatibility settings, not a verified support claim.
Actual Server 2019 and Windows 11 acceptance is still required. Citrix/RDP
remote-session integration is not implemented.

## What is implemented

- One descriptorless virtual monitor, 1920 × 1080, 60 Hz, SDR; no borrowed EDID.
- A Direct3D frame consumer, cancellation-aware processing thread and explicit
  swapchain ownership. No frame encoding, streaming, audio or HDR.
- A versioned, read-only status IOCTL reporting adapter initialization, monitor
  arrival, active mode, power state, swapchain state and completed frames.
- `DesktopVddCtl.exe hold`: creates one software device and holds its lifetime
  handle until the requested duration ends or Ctrl+C is received.
- `DesktopVddCtl.exe status`: queries the driver; it does not install or create it.
- Owned-device removal observation after a normal hold, without uninstalling the
  staged package. Process termination closes the handle; PnP removal is asynchronous.
- WinForms and WPF fixtures with stable AutomationIds, in-memory records, numeric
  validation and a changing screenshot marker.
- A pywinauto acceptance script with JSON/JUnit results and screenshot evidence.

The privileged service/broker, user-session worker, generic supervised `run`
command, automatic session provisioning and production installer are subsequent
increments. The utility does not launch test code as administrator. Keep the
fixture and your automation engine in the intended test-user session.

## Build

Use Windows with Visual Studio 2022 C++ build tools (including Spectre libraries),
PowerShell 7, NuGet CLI and .NET 8 SDK. The script restores pinned SDK/WDK NuGet
packages `10.0.26100.6584`; the matching kit build is `26100`. No driver is
installed on the build machine. Do not use `dotnet build` for the native projects.

From the repository root:

```powershell
pwsh ./scripts/Build-VddPrototype.ps1
```

Output: `artifacts/vdd/package/`. The separate **VDD Prototype** workflow builds
the same package on Windows and uploads
`DesktopVdd-prototype-unsigned-win-x64`. Its `build-evidence.json` explicitly
leaves installation and unattended execution unverified. The existing solution
and driver/agent release packages remain separate.

The build also runs portable/native contract tests. To run just these on Linux:

```bash
g++ -std=c++17 -Wall -Wextra -Werror src/DesktopVirtualDisplay/Tests/Contracts.cpp -o /tmp/vdd-contracts
/tmp/vdd-contracts
```

## Signing and staging

Obtain a signing route approved for the target machine policy. No signing key or
certificate is included in this repository. Do not treat a publisher signature
as Microsoft certification or a guarantee of corporate acceptance.

The optional signing script takes an existing certificate with an accessible
private key. It signs the DLL first, regenerates the catalog, signs the catalog,
then verifies membership and local trust. Supply the actual SDK/WDK tool paths:

```powershell
pwsh ./scripts/Sign-VddPrototype.ps1 `
  -PackageDirectory ./artifacts/vdd/package `
  -CertificateThumbprint '<approved-certificate-thumbprint>' `
  -SignToolPath '<path-to-signtool.exe>' `
  -Inf2CatPath '<path-to-Inf2Cat.exe>'
```

An approved HTTPS timestamp URL can be supplied with `-TimestampUrl`. The
script neither creates a certificate nor changes machine trust or boot policy.
Re-sign/rebuild the catalog after changing any INF-referenced package file.

On an isolated test machine with the required trust/approval, open elevated
PowerShell and stage the signed package:

```powershell
./Install-VddPrototype.ps1 -PackageDirectory .
```

The installer refuses an untrusted/unsigned catalog, then invokes PnPUtil staging.
Windows still checks the whole package and policy. Record the resulting `oemNN.inf`
name for this exact package. Staging alone creates no permanent display device.
The package contains a native DLL/INF/catalog; IddCx itself is supplied by Windows.

## First monitor test

Use a dedicated, recoverable Windows machine/VM with an already provisioned
console user session. Start with a VM console viewer so failures can be observed.
From an elevated terminal, in the package directory:

```powershell
./DesktopVddCtl.exe hold --seconds 300 --startup-timeout 30
```

The fixed software-device identity admits only one active instance. The utility
never removes another process's device and does not adopt a device by its name.
`hold` defaults to 300 seconds, allows 1–3600, and waits up to the configured
timeout separately for enumeration and pipeline readiness. Ctrl+C cancels.

Success emits `display-pipeline-ready` only after an active 1080p path, swapchain
and completed frame are observed. If Windows has left the monitor disconnected,
activate it through Windows Display Settings in this isolated lab and repeat.
This version does not silently rewrite display topology, primary monitor, DPI
or the settings of another monitor. Check 100% test-user scaling separately.

While `hold` remains alive, a second elevated terminal can query:

```powershell
./DesktopVddCtl.exe status
```

Output is newline-delimited JSON. `pipelineReady` is **not** input/capture/test
readiness; those fields remain false. Status counters are an observational
snapshot, not atomic proof that the state cannot change immediately afterward.
`lastDriverError` records the last diagnostic HRESULT/NTSTATUS and can persist
after the OS retries a swapchain successfully.

| Exit | Meaning |
| --- | --- |
| 0 | Status sees an active pipeline, or hold completed and observed device removal |
| 2 | Device/pipeline unavailable, permission failure or infrastructure error |
| 3 | Hold completed, but removal was not observed within ten seconds |
| 64 | Invalid syntax; no device creation attempted |
| 130 | Ctrl+C cancellation; cleanup is still attempted |

If startup fails before a successful creation callback, removal cannot be
positively identified and is reported as unknown. Closing the lifetime handle
still requests removal. A query timeout cancels its I/O and waits for completion
to keep native buffers alive; severe OS/driver hangs can still prevent a bounded
process exit. No service watchdog is supplied in this increment.

## Test real UI operations

Run the fixture and automation in the same eligible user session and integrity
context. Keep the elevated display holder separate. A Jenkins service launching
an ordinary child in Session 0 is not a substitute for a user-session worker.

Find active display names in a PowerShell process inside that user session:

```powershell
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Screen]::AllScreens | Select-Object DeviceName, Bounds, Primary
```

Identify the new monitor in Windows Display Settings; do not assume it is always
`DISPLAY2`. Substitute its actual name below. The holder must remain active for
the entire test; increase its duration when testing manually.

```powershell
python -m pip install -r ./acceptance/requirements.txt
python ./acceptance/fixture_smoke.py `
  --fixture ./fixtures/WinForms/DesktopRuntime.Fixtures.WinForms.exe `
  --display '\\.\DISPLAY2' --output ./results/winforms

python ./acceptance/fixture_smoke.py `
  --fixture ./fixtures/Wpf/DesktopRuntime.Fixtures.Wpf.exe `
  --display '\\.\DISPLAY2' --output ./results/wpf
```

The script exercises real mouse/keyboard input, required-name and numeric-range
validation, a disabled field, dropdown selection, add/grid assertion, update/grid
assertion, deletion and absence. It compares fresh crops of the changing marker
and saves the full window. It does not retry uncertain mutations. All fixture
records are in memory; cleanup terminates only the launched fixture process.

Run once with `--expect-wrong-row`: it must return exit 1 and write a failed JUnit
result. A negative-control failure is expected, never relabelled as a passing
application test. Exit 2 indicates an infrastructure error and exit 3 a cleanup
failure. Archive `acceptance.json`, `junit.xml` and PNGs even when the run fails.
This acceptance harness still needs execution against both actual GUI frameworks;
CI syntax/build checks do not establish its UI locator behavior.

For DesktopAutomationHelper, use the same fixture AutomationIds through its
existing locator/action APIs. This increment does not introduce a new JSON plan
format or claim a validated end-to-end engine adapter:

| AutomationId | Expected use |
| --- | --- |
| `NameInput`, `AgeInput` | Enabled text fields; type and verify value |
| `RoleInput` | Reader / Editor / Administrator dropdown |
| `DisabledInput` | Must be disabled |
| `AddRecord`, `UpdateRecord`, `DeleteRecord` | CRUD actions |
| `RecordsGrid` | Row identity, count, value and absence assertions |
| `ResultText` | Validation/success message |
| `FrameMarker` | Changing capture evidence |

## No-viewer acceptance gate

After the visible-console baseline, close every viewer and run the same test
from an independently established eligible user-session process. Record how the
session was provisioned and its Windows session ID. Do not use a minimized or
disconnected RDP success as interchangeable evidence for console execution.

Collect separate outcomes for visible RDP, minimized RDP, disconnected RDP,
console without viewer, locked/logged-off negative cases, application/holder
crashes, repeated acquire/release and concurrent acquire rejection. Confirm the
owned device disappears after hold termination, and that the package remains
available for the next run. Monitor/lock transitions should never become a false
functional pass. Use 30 nominal repetitions per supported profile before making
a release reliability claim, plus deliberate failure cases.

Server 2019's IddCx 1.3 predates remote-session IDDs. A console display can neither
create a Citrix session nor become a per-Citrix-session display through a flag.
The first test result must explicitly identify which session owns the application.

## Cleanup and uninstall

Ctrl+C the holder or let its duration finish. Check its cleanup event, Device
Manager and `status`; after removal, status should return unavailable (exit 2).
Do not use `pnputil /remove-device` against an unrelated or guessed device ID.

To uninstall the staged package, first stop all holders, then inspect:

```powershell
pnputil /enum-drivers
# Only after matching the recorded OEM name, provider, original INF and version:
pnputil /delete-driver oemNN.inf
```

Do not use forced deletion of an in-use package. Preserve SetupAPI/PnP error
output and restore the test VM snapshot if recovery is needed. Corporate
deployment, reboot recovery and upgrades require further testing.

## References

- [Microsoft IDD model](https://learn.microsoft.com/en-us/windows-hardware/drivers/display/indirect-display-driver-model-overview)
- [IddCx OS versions](https://learn.microsoft.com/en-us/windows-hardware/drivers/display/iddcx-versions)
- [IddCx 1.3 compatibility](https://learn.microsoft.com/en-us/windows-hardware/drivers/display/building-iddcx1.4-drivers)
- [Software-device creation](https://learn.microsoft.com/en-us/windows/win32/api/swdevice/nf-swdevice-swdevicecreate)
- [Software-device lifetime](https://learn.microsoft.com/en-us/windows/win32/api/swdevice/nf-swdevice-swdevicesetlifetime)
- [Frame ownership/error handling](https://learn.microsoft.com/en-us/windows-hardware/drivers/display/idd-evtiddcxmonitorassignswapchain-error-handling)
- [Build-tool package guidance](https://learn.microsoft.com/en-us/windows-hardware/drivers/install-the-wdk-using-nuget)

The native sample attribution and license are included in every package.
