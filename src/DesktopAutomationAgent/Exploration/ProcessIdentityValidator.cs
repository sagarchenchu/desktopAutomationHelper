using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using DesktopAutomationAgent.Observation;

namespace DesktopAutomationAgent.Exploration;

public sealed class ProcessIdentityInfo
{
    public int ProcessId { get; init; }

    public string? ProcessName { get; init; }

    public string? ExecutablePath { get; init; }

    public string? ExecutableSha256 { get; init; }

    public int? SessionId { get; init; }

    public int? ParentProcessId { get; init; }
}

public interface IProcessIdentityProbe
{
    bool TryGetIdentity(int processId, out ProcessIdentityInfo info, out string? error);
}

/// <summary>
/// Resolves PID → process name / path / hash / session / parent via System.Diagnostics
/// and the supported Toolhelp32 snapshot API for parent lineage.
/// </summary>
public sealed class SystemProcessIdentityProbe : IProcessIdentityProbe
{
    public bool TryGetIdentity(int processId, out ProcessIdentityInfo info, out string? error)
    {
        info = new ProcessIdentityInfo { ProcessId = processId };
        error = null;

        if (processId <= 0)
        {
            error = "Process id is missing or invalid.";
            return false;
        }

        try
        {
            using var process = Process.GetProcessById(processId);
            string? path = null;
            string? sha = null;
            string? pathError = null;
            try
            {
                path = process.MainModule?.FileName;
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    var bytes = SHA256.HashData(File.ReadAllBytes(path));
                    sha = Convert.ToHexString(bytes).ToLowerInvariant();
                }
            }
            catch (Exception ex)
            {
                pathError = $"Unable to read executable for PID {processId}: {ex.GetType().Name}.";
            }

            int? sessionId = null;
            try
            {
                sessionId = process.SessionId;
            }
            catch
            {
                // Session may be inaccessible for some processes.
            }

            if (!ToolhelpParentProcessId.TryGet(processId, out var parentPid, out var parentError))
            {
                error = parentError ?? $"Unable to resolve parent process id for PID {processId}.";
                return false;
            }

            info = new ProcessIdentityInfo
            {
                ProcessId = processId,
                ProcessName = string.IsNullOrWhiteSpace(process.ProcessName) ? null : process.ProcessName,
                ExecutablePath = path,
                ExecutableSha256 = sha,
                SessionId = sessionId,
                ParentProcessId = parentPid
            };

            if (string.IsNullOrWhiteSpace(info.ProcessName))
            {
                error = $"Process name missing for PID {processId}.";
                return false;
            }

            if (pathError is not null && string.IsNullOrWhiteSpace(path))
                error = pathError;

            return true;
        }
        catch (Exception ex)
        {
            error = $"Unable to resolve process identity for PID {processId}: {ex.GetType().Name}.";
            return false;
        }
    }
}

/// <summary>
/// Parent PID via CreateToolhelp32Snapshot / Process32FirstW / Process32NextW (supported Win32 API).
/// </summary>
internal static class ToolhelpParentProcessId
{
    private const uint Th32csSnapProcess = 0x00000002;

    public static bool TryGet(int processId, out int parentProcessId, out string? error)
    {
        parentProcessId = 0;
        error = null;

        if (!OperatingSystem.IsWindows())
        {
            error = "Parent process lookup requires Windows.";
            return false;
        }

        var snapshot = CreateToolhelp32Snapshot(Th32csSnapProcess, 0);
        if (snapshot == IntPtr.Zero || snapshot == new IntPtr(-1))
        {
            error = $"CreateToolhelp32Snapshot failed for PID {processId}.";
            return false;
        }

        try
        {
            var entry = new Processentry32W
            {
                DwSize = (uint)Marshal.SizeOf<Processentry32W>()
            };

            if (!Process32FirstW(snapshot, ref entry))
            {
                error = $"Process32FirstW failed for PID {processId}.";
                return false;
            }

            do
            {
                if (unchecked((int)entry.Th32ProcessId) == processId)
                {
                    parentProcessId = unchecked((int)entry.Th32ParentProcessId);
                    if (parentProcessId <= 0)
                    {
                        error = $"Parent process id missing for PID {processId}.";
                        return false;
                    }

                    return true;
                }
            } while (Process32NextW(snapshot, ref entry));

            error = $"PID {processId} not found in Toolhelp32 snapshot.";
            return false;
        }
        finally
        {
            CloseHandle(snapshot);
        }
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessId);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool Process32FirstW(IntPtr hSnapshot, ref Processentry32W lppe);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool Process32NextW(IntPtr hSnapshot, ref Processentry32W lppe);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct Processentry32W
    {
        public uint DwSize;
        public uint CntUsage;
        public uint Th32ProcessId;
        public UIntPtr Th32DefaultHeapId;
        public uint Th32ModuleId;
        public uint CntThreads;
        public uint Th32ParentProcessId;
        public int PcPriClassBase;
        public uint DwFlags;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string SzExeFile;
    }
}

public sealed class ProcessIdentityValidationResult
{
    public bool IsValid { get; init; }

    public string? Detail { get; init; }
}

/// <summary>
/// Fail-closed application identity checks for a single exploration crawl.
/// Root PID/session are pinned once; root windows must match that pin; child
/// ancestry must terminate at the pinned PID (no process-name shortcut).
/// Create a new instance per ExploreAsync — do not share across crawls.
/// </summary>
public sealed class ProcessIdentityValidator
{
    private readonly IProcessIdentityProbe _probe;
    private int? _pinnedRootProcessId;
    private int? _pinnedRootSessionId;
    private bool _rootPinned;

    public ProcessIdentityValidator(IProcessIdentityProbe? probe = null)
    {
        _probe = probe ?? new SystemProcessIdentityProbe();
    }

    public int? PinnedRootProcessId => _pinnedRootProcessId;

    public int? PinnedRootSessionId => _pinnedRootSessionId;

    public bool IsRootPinned => _rootPinned;

    /// <summary>
    /// Pin the initial in-scope root process for the crawl. Subsequent calls are no-ops.
    /// Live session id is required (fail-closed).
    /// </summary>
    public void PinRoot(int processId, int? sessionId)
    {
        if (_rootPinned)
            return;

        if (processId <= 0)
            throw new ArgumentOutOfRangeException(nameof(processId));
        if (sessionId is null or <= 0)
            throw new ArgumentOutOfRangeException(nameof(sessionId), "Live session id is required to pin root.");

        _pinnedRootProcessId = processId;
        _pinnedRootSessionId = sessionId;
        _rootPinned = true;
    }

    public ProcessIdentityValidationResult ValidateSnapshot(
        SemanticSnapshot snapshot,
        ExplorationPolicyDocument policy)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(policy);

        if (string.IsNullOrWhiteSpace(snapshot.Application.ProcessName))
            return Fail("Observed application processName is missing.");

        if (!IsProcessNameAllowed(snapshot.Application.ProcessName, policy))
        {
            return Fail(
                $"Process '{snapshot.Application.ProcessName}' is outside the allowed process tree.");
        }

        if (snapshot.Application.ProcessId is not int pid || pid <= 0)
            return Fail("Observed application processId is missing.");

        if (!_probe.TryGetIdentity(pid, out var identity, out var probeError))
            return Fail(probeError ?? $"Unable to resolve process identity for PID {pid}.");

        if (!string.Equals(
                identity.ProcessName,
                snapshot.Application.ProcessName,
                StringComparison.OrdinalIgnoreCase))
        {
            return Fail(
                $"Live process name '{identity.ProcessName}' does not match snapshot '{snapshot.Application.ProcessName}'.");
        }

        var pathHash = ValidatePathAndHash(identity, policy);
        if (!pathHash.IsValid)
            return pathHash;

        // Live probed session is authoritative — never let snapshot override it.
        if (identity.SessionId is not int liveSession)
            return Fail($"Live Windows session id missing for PID {pid}; fail-closed.");

        if (snapshot.WindowsSession.SessionId is int snapSession && snapSession != liveSession)
        {
            return Fail(
                $"Snapshot session {snapSession} disagrees with live session {liveSession}; fail-closed.");
        }

        if (policy.Scope.WindowsSessionId is int expectedSession && liveSession != expectedSession)
        {
            return Fail(
                $"Windows session {liveSession} differs from allowed session {expectedSession}.");
        }

        if (!_rootPinned)
        {
            // First successful observation pins the crawl root (live PID + live session).
            if (!string.Equals(
                    snapshot.Application.ProcessName,
                    policy.Application.ProcessName,
                    StringComparison.OrdinalIgnoreCase))
            {
                return Fail(
                    $"Initial process '{snapshot.Application.ProcessName}' is not the policy root '{policy.Application.ProcessName}'.");
            }

            PinRoot(pid, liveSession);
            return Ok();
        }

        var sessionPin = EnforcePinnedSession(liveSession);
        if (!sessionPin.IsValid)
            return sessionPin;

        // After pin: root windows must equal pinned PID + pinned session.
        if (pid == _pinnedRootProcessId)
            return Ok();

        // Different PID — must be an allowed child whose ancestry terminates at the pinned root.
        return ValidateChildLineage(pid, identity, policy);
    }

    public ProcessIdentityValidationResult ValidateNewWindow(
        WindowDescriptor window,
        ExplorationPolicyDocument policy)
    {
        ArgumentNullException.ThrowIfNull(window);
        ArgumentNullException.ThrowIfNull(policy);

        if (string.IsNullOrWhiteSpace(window.ProcessName))
            return Fail($"New window '{window.Title}' is missing processName; fail-closed.");

        if (window.ProcessId is not int pid || pid <= 0)
            return Fail($"New window '{window.Title}' is missing processId; fail-closed.");

        if (!_probe.TryGetIdentity(pid, out var identity, out var probeError))
            return Fail(probeError ?? $"Unable to resolve process identity for new window PID {pid}.");

        if (!string.Equals(identity.ProcessName, window.ProcessName, StringComparison.OrdinalIgnoreCase))
        {
            return Fail(
                $"New window process name mismatch: listed '{window.ProcessName}', live '{identity.ProcessName}'.");
        }

        if (!_rootPinned || _pinnedRootProcessId is not int rootPid)
            return Fail("Root process is not pinned; cannot validate new window.");

        if (identity.SessionId is not int liveSession)
            return Fail($"Live Windows session id missing for new window PID {pid}; fail-closed.");

        var sessionPin = EnforcePinnedSession(liveSession);
        if (!sessionPin.IsValid)
            return sessionPin;

        // Same process as pinned root — session already enforced.
        if (pid == rootPid)
        {
            var pathHash = ValidatePathAndHash(identity, policy);
            return pathHash.IsValid ? Ok() : pathHash;
        }

        // Different PID: no process-name==root shortcut — must be allowed child + ancestry to pin.
        if (!policy.Application.AllowedChildProcessNames.Any(c =>
                string.Equals(c, window.ProcessName, StringComparison.OrdinalIgnoreCase)))
        {
            return Fail(
                $"Unexpected external window '{window.Title}' / process '{window.ProcessName}' (pid {pid}).");
        }

        return ValidateChildLineage(pid, identity, policy);
    }

    /// <summary>
    /// Verifies a live HWND's process id matches the expected PID after switchwindow
    /// or against a pre-action/initial window inventory.
    /// </summary>
    public ProcessIdentityValidationResult ValidateHwndProcess(
        long hwnd,
        int expectedProcessId,
        WindowInventory inventory)
    {
        ArgumentNullException.ThrowIfNull(inventory);

        var match = inventory.Windows.FirstOrDefault(w => w.Hwnd == hwnd);
        if (match is null)
            return Fail($"HWND 0x{hwnd:X} is not present in the window inventory.");

        if (match.ProcessId is not int pid || pid <= 0)
            return Fail($"HWND 0x{hwnd:X} is missing processId; fail-closed.");

        if (pid != expectedProcessId)
        {
            return Fail(
                $"HWND 0x{hwnd:X} processId {pid} does not match expected {expectedProcessId}; fail-closed.");
        }

        return Ok();
    }

    /// <summary>
    /// Validates the active snapshot window's ProcessId against the pinned process tree
    /// (root or allowed-child lineage). Does not accept Application.ProcessId as a
    /// substitute for Window.ProcessId.
    /// </summary>
    public ProcessIdentityValidationResult ValidateActiveWindow(
        WindowIdentity window,
        ExplorationPolicyDocument policy)
    {
        ArgumentNullException.ThrowIfNull(window);
        ArgumentNullException.ThrowIfNull(policy);

        if (window.Hwnd is not long hwnd || hwnd <= 0)
            return Fail("Active window HWND is missing; fail-closed.");

        if (window.ProcessId is not int pid || pid <= 0)
            return Fail("Active window processId is missing; fail-closed.");

        if (!_probe.TryGetIdentity(pid, out var identity, out var probeError))
            return Fail(probeError ?? $"Unable to resolve process identity for active window PID {pid}.");

        if (string.IsNullOrWhiteSpace(identity.ProcessName))
            return Fail($"Active window PID {pid} is missing a live process name; fail-closed.");

        // Reuse new-window rules: same session pin + root or allowed-child ancestry.
        return ValidateNewWindow(
            new WindowDescriptor
            {
                Hwnd = hwnd,
                Title = window.Title,
                ClassName = window.ClassName,
                ProcessId = pid,
                ProcessName = identity.ProcessName
            },
            policy);
    }

    private ProcessIdentityValidationResult ValidatePathAndHash(
        ProcessIdentityInfo identity,
        ExplorationPolicyDocument policy)
    {
        var requiresPath = !string.IsNullOrWhiteSpace(policy.Application.ExecutablePath);
        var requiresHash = !string.IsNullOrWhiteSpace(policy.Application.ExecutableSha256);

        if (!requiresPath && !requiresHash)
            return Ok();

        // Path/hash constraints apply only to the pinned root process identity.
        if (_rootPinned
            && _pinnedRootProcessId is int rootPid
            && identity.ProcessId != rootPid)
        {
            return Ok();
        }

        if (requiresPath)
        {
            if (string.IsNullOrWhiteSpace(identity.ExecutablePath))
                return Fail("Executable path could not be resolved; fail-closed.");

            if (!PathsEqual(identity.ExecutablePath, policy.Application.ExecutablePath!))
            {
                return Fail(
                    $"Executable path '{identity.ExecutablePath}' does not match policy '{policy.Application.ExecutablePath}'.");
            }
        }

        if (requiresHash)
        {
            if (string.IsNullOrWhiteSpace(identity.ExecutableSha256))
                return Fail("Executable SHA256 could not be resolved; fail-closed.");

            if (!string.Equals(
                    identity.ExecutableSha256,
                    policy.Application.ExecutableSha256,
                    StringComparison.OrdinalIgnoreCase))
            {
                return Fail("Executable SHA256 does not match policy.");
            }
        }

        return Ok();
    }

    private ProcessIdentityValidationResult ValidateChildLineage(
        int childPid,
        ProcessIdentityInfo identity,
        ExplorationPolicyDocument policy)
    {
        if (!policy.Application.AllowedChildProcessNames.Any(c =>
                string.Equals(c, identity.ProcessName, StringComparison.OrdinalIgnoreCase)))
        {
            return Fail($"Process '{identity.ProcessName}' is not an allowed child.");
        }

        if (_pinnedRootProcessId is not int rootPid)
            return Fail("Root process is not pinned; cannot verify child lineage.");

        if (identity.SessionId is not int childSession)
            return Fail($"Live Windows session id missing for child PID {childPid}; fail-closed.");

        var sessionPin = EnforcePinnedSession(childSession);
        if (!sessionPin.IsValid)
            return sessionPin;

        // Walk parent chain until it terminates at the pinned root PID.
        // Do not accept process-name equality as a substitute for PID ancestry.
        var current = identity;
        var seen = new HashSet<int> { childPid };

        for (var hops = 0; hops < 32; hops++)
        {
            if (current.ParentProcessId is not int parent || parent <= 0)
                return Fail($"Parent process id missing while walking lineage from PID {childPid}; fail-closed.");

            if (parent == rootPid)
                return Ok();

            if (!seen.Add(parent))
                return Fail($"Cycle detected in process lineage for PID {childPid}; fail-closed.");

            if (!_probe.TryGetIdentity(parent, out var parentInfo, out var parentError))
                return Fail(parentError ?? $"Unable to resolve parent PID {parent} for child {childPid}.");

            if (parentInfo.SessionId is int parentSession)
            {
                var parentSessionPin = EnforcePinnedSession(parentSession);
                if (!parentSessionPin.IsValid)
                    return parentSessionPin;
            }

            if (parentInfo.ProcessId == rootPid)
                return Ok();

            current = parentInfo;
        }

        return Fail(
            $"Child process '{identity.ProcessName}' (pid {childPid}) ancestry did not reach pinned root {rootPid}.");
    }

    private ProcessIdentityValidationResult EnforcePinnedSession(int liveSession)
    {
        if (_pinnedRootSessionId is not int pinnedSession)
            return Fail("Pinned Windows session id is missing; fail-closed.");

        if (liveSession != pinnedSession)
        {
            return Fail(
                $"Live session {liveSession} differs from pinned session {pinnedSession}; fail-closed.");
        }

        return Ok();
    }

    private static bool IsProcessNameAllowed(string processName, ExplorationPolicyDocument policy)
    {
        if (string.Equals(processName, policy.Application.ProcessName, StringComparison.OrdinalIgnoreCase))
            return true;

        return policy.Application.AllowedChildProcessNames
            .Any(child => string.Equals(child, processName, StringComparison.OrdinalIgnoreCase));
    }

    private static bool PathsEqual(string left, string right)
    {
        try
        {
            return string.Equals(
                Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
        }
        catch
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static ProcessIdentityValidationResult Ok() => new() { IsValid = true };

    private static ProcessIdentityValidationResult Fail(string detail) =>
        new() { IsValid = false, Detail = detail };
}
