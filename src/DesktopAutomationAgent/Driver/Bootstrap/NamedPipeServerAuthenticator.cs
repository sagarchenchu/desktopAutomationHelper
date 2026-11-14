using System.ComponentModel;
using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;

namespace DesktopAutomationAgent.Driver.Bootstrap;

/// <summary>
/// Verifies the remote named-pipe server process identity via OS APIs before
/// trusting bootstrap payload fields (SID / PID / token).
/// </summary>
public static class NamedPipeServerAuthenticator
{
    /// <summary>
    /// After <see cref="NamedPipeClientStream.ConnectAsync"/>, resolve the server
    /// process id and its Windows user SID. Returns false on any OS failure.
    /// </summary>
    [SupportedOSPlatform("windows")]
    public static bool TryGetServerIdentity(
        NamedPipeClientStream client,
        out int serverProcessId,
        out string? serverUserSid)
    {
        serverProcessId = 0;
        serverUserSid = null;

        if (!OperatingSystem.IsWindows())
            return false;

        ArgumentNullException.ThrowIfNull(client);

        if (!GetNamedPipeServerProcessId(client.SafePipeHandle, out var pid))
            return false;

        if (pid == 0 || pid > int.MaxValue)
            return false;

        serverProcessId = (int)pid;
        serverUserSid = TryGetProcessUserSid(serverProcessId);
        return !string.IsNullOrWhiteSpace(serverUserSid);
    }

    [SupportedOSPlatform("windows")]
    public static string? TryGetProcessUserSid(int processId)
    {
        if (processId <= 0)
            return null;

        SafeProcessHandle? processHandle = null;
        SafeAccessTokenHandle? tokenHandle = null;
        try
        {
            processHandle = OpenProcess(ProcessAccessFlags.QueryLimitedInformation, false, processId);
            if (processHandle.IsInvalid)
                return null;

            if (!OpenProcessToken(processHandle, TokenAccessLevels.Query, out tokenHandle) || tokenHandle.IsInvalid)
                return null;

            using var identity = new WindowsIdentity(tokenHandle.DangerousGetHandle());
            return identity.User?.Value;
        }
        catch (ArgumentException)
        {
            return null;
        }
        catch (Win32Exception)
        {
            return null;
        }
        finally
        {
            tokenHandle?.Dispose();
            processHandle?.Dispose();
        }
    }

    [SupportedOSPlatform("windows")]
    public static DateTimeOffset? TryGetProcessStartTimeUtc(int processId) =>
        StaleBootstrapDetector.TryGetProcessStartTimeUtc(processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetNamedPipeServerProcessId(
        SafePipeHandle pipe,
        out uint serverProcessId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern SafeProcessHandle OpenProcess(
        ProcessAccessFlags processAccess,
        bool bInheritHandle,
        int processId);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool OpenProcessToken(
        SafeProcessHandle processHandle,
        TokenAccessLevels desiredAccess,
        out SafeAccessTokenHandle tokenHandle);

    [Flags]
    private enum ProcessAccessFlags : uint
    {
        QueryLimitedInformation = 0x1000
    }

    [Flags]
    private enum TokenAccessLevels : uint
    {
        Query = 0x0008
    }
}
