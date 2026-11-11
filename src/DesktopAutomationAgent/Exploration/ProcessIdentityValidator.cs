using System.Diagnostics;
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
/// Resolves PID → process name / path / hash / session via System.Diagnostics.
/// Fail-closed callers treat probe failures as out of scope.
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
                error = $"Unable to read executable for PID {processId}: {ex.GetType().Name}.";
                // Continue with name/session when path is unavailable; callers decide fail-closed.
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

            info = new ProcessIdentityInfo
            {
                ProcessId = processId,
                ProcessName = string.IsNullOrWhiteSpace(process.ProcessName) ? null : process.ProcessName,
                ExecutablePath = path,
                ExecutableSha256 = sha,
                SessionId = sessionId,
                // Parent PID is optional; tests inject probes with lineage. Live probe leaves
                // it null when undocumented NT APIs are unavailable — name allow-list still applies.
                ParentProcessId = null
            };

            if (string.IsNullOrWhiteSpace(info.ProcessName))
            {
                error = $"Process name missing for PID {processId}.";
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            error = $"Unable to resolve process identity for PID {processId}: {ex.GetType().Name}.";
            return false;
        }
    }
}

public sealed class ProcessIdentityValidationResult
{
    public bool IsValid { get; init; }

    public string? Detail { get; init; }
}

/// <summary>
/// Fail-closed application identity checks: process name, optional path/SHA256,
/// Windows session, and parent/child lineage for allowed children.
/// </summary>
public sealed class ProcessIdentityValidator
{
    private readonly IProcessIdentityProbe _probe;
    private int? _rootProcessId;

    public ProcessIdentityValidator(IProcessIdentityProbe? probe = null)
    {
        _probe = probe ?? new SystemProcessIdentityProbe();
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

        if (policy.Scope.WindowsSessionId is int expectedSession)
        {
            var actualSession = snapshot.WindowsSession.SessionId ?? identity.SessionId;
            if (actualSession is null)
                return Fail("Windows session id is missing; fail-closed.");
            if (actualSession != expectedSession)
            {
                return Fail(
                    $"Windows session {actualSession} differs from allowed session {expectedSession}.");
            }
        }

        var isRoot = string.Equals(
            snapshot.Application.ProcessName,
            policy.Application.ProcessName,
            StringComparison.OrdinalIgnoreCase);

        if (isRoot)
        {
            _rootProcessId = pid;
        }
        else
        {
            var lineage = ValidateChildLineage(pid, identity, policy);
            if (!lineage.IsValid)
                return lineage;
        }

        return Ok();
    }

    public ProcessIdentityValidationResult ValidateNewWindow(
        WindowDescriptor window,
        ExplorationPolicyDocument policy)
    {
        ArgumentNullException.ThrowIfNull(window);
        ArgumentNullException.ThrowIfNull(policy);

        if (string.IsNullOrWhiteSpace(window.ProcessName))
            return Fail($"New window '{window.Title}' is missing processName; fail-closed.");

        if (!IsProcessNameAllowed(window.ProcessName, policy))
        {
            return Fail(
                $"Unexpected external window '{window.Title}' / process '{window.ProcessName}'.");
        }

        if (window.ProcessId is not int pid || pid <= 0)
            return Fail($"New window '{window.Title}' is missing processId; fail-closed.");

        var requiresPathOrHash = !string.IsNullOrWhiteSpace(policy.Application.ExecutablePath)
                                 || !string.IsNullOrWhiteSpace(policy.Application.ExecutableSha256);

        if (_probe.TryGetIdentity(pid, out var identity, out var probeError))
        {
            if (!string.Equals(identity.ProcessName, window.ProcessName, StringComparison.OrdinalIgnoreCase))
            {
                return Fail(
                    $"New window process name mismatch: listed '{window.ProcessName}', live '{identity.ProcessName}'.");
            }

            var pathHash = ValidatePathAndHash(identity, policy, allowChildWithoutRootHash: true);
            if (!pathHash.IsValid)
                return pathHash;

            if (!string.Equals(
                    window.ProcessName,
                    policy.Application.ProcessName,
                    StringComparison.OrdinalIgnoreCase))
            {
                var lineage = ValidateChildLineage(pid, identity, policy);
                if (!lineage.IsValid)
                    return lineage;
            }

            return Ok();
        }

        // Probe failed: fail closed when path/hash must be enforced; otherwise processName
        // from listwindows is the identity gate for new windows.
        if (requiresPathOrHash
            && string.Equals(
                window.ProcessName,
                policy.Application.ProcessName,
                StringComparison.OrdinalIgnoreCase))
        {
            return Fail(probeError ?? $"Unable to resolve process identity for new window PID {pid}.");
        }

        return Ok();
    }

    private ProcessIdentityValidationResult ValidatePathAndHash(
        ProcessIdentityInfo identity,
        ExplorationPolicyDocument policy,
        bool allowChildWithoutRootHash = false)
    {
        var requiresPath = !string.IsNullOrWhiteSpace(policy.Application.ExecutablePath);
        var requiresHash = !string.IsNullOrWhiteSpace(policy.Application.ExecutableSha256);

        var isRoot = string.Equals(
            identity.ProcessName,
            policy.Application.ProcessName,
            StringComparison.OrdinalIgnoreCase);

        if (!isRoot && allowChildWithoutRootHash && !requiresPath && !requiresHash)
            return Ok();

        // Path/hash constraints apply to the root application identity. Allowed children
        // are gated by name + lineage; when set on policy they still require resolvable identity.
        if (requiresPath || requiresHash)
        {
            if (!isRoot)
                return Ok();

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

        if (_rootProcessId is int rootPid)
        {
            if (identity.ParentProcessId is int parent)
            {
                if (parent != rootPid && parent != childPid)
                {
                    // Walk one level: accept if parent is root; otherwise fail closed when known.
                    if (!_probe.TryGetIdentity(parent, out var parentInfo, out _))
                        return Fail($"Unable to verify parent lineage for child PID {childPid}.");

                    if (parentInfo.ProcessId != rootPid
                        && !string.Equals(
                            parentInfo.ProcessName,
                            policy.Application.ProcessName,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return Fail(
                            $"Child process '{identity.ProcessName}' parent '{parentInfo.ProcessName}' is outside lineage.");
                    }
                }
            }
            // Missing parent id: name allow-list already checked; do not invent lineage.
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
