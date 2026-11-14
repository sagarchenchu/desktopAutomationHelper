using System.Diagnostics;

namespace DesktopAutomationAgent.Driver.Bootstrap;

public static class StaleBootstrapDetector
{
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

        if (string.IsNullOrWhiteSpace(payload.Token) && string.IsNullOrWhiteSpace(payload.AuthorizationHeader))
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
