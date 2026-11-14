using System.Diagnostics;

namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// Detects stale bootstrap descriptors left behind after a driver crash.
/// </summary>
public static class StaleBootstrapDetector
{
    /// <summary>
    /// Returns true when the descriptor should be ignored / cleaned up because
    /// the owning process is no longer alive (or the payload is incomplete).
    /// </summary>
    public static bool IsStale(SecureBootstrapPayload? payload, Func<int, bool>? isProcessAlive = null)
    {
        if (payload is null)
            return true;

        if (payload.SchemaVersion != SecureBootstrapConstants.SchemaVersion)
            return true;

        if (string.IsNullOrWhiteSpace(payload.UserSid))
            return true;

        if (payload.Port <= 0 || payload.Pid <= 0)
            return true;

        if (string.IsNullOrWhiteSpace(payload.Token))
            return true;

        var alive = isProcessAlive ?? IsProcessAlive;
        return !alive(payload.Pid);
    }

    public static bool IsProcessAlive(int pid)
    {
        if (pid <= 0)
            return false;

        try
        {
            using var process = Process.GetProcessById(pid);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
    }
}
