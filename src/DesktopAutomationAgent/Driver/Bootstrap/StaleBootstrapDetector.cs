using System.Diagnostics;

namespace DesktopAutomationAgent.Driver.Bootstrap;

public static class StaleBootstrapDetector
{
    public static bool IsStale(
        SecureBootstrapPayload? payload,
        Func<int, DateTimeOffset?>? getProcessStartTimeUtc = null)
    {
        if (payload is null)
            return true;

        if (payload.SchemaVersion != SecureBootstrapConstants.SchemaVersion)
            return true;

        if (string.IsNullOrWhiteSpace(payload.UserSid))
            return true;

        if (payload.Port <= 0 || payload.Pid <= 0)
            return true;

        if (payload.ProcessStartTimeUtc == default)
            return true;

        if (string.IsNullOrWhiteSpace(payload.Token) && string.IsNullOrWhiteSpace(payload.AuthorizationHeader))
            return true;

        var lookup = getProcessStartTimeUtc ?? TryGetProcessStartTimeUtc;
        var actual = lookup(payload.Pid);
        if (actual is null)
            return true;

        var delta = (actual.Value - payload.ProcessStartTimeUtc).Duration();
        return delta > SecureBootstrapConstants.ProcessStartTimeTolerance;
    }

    public static DateTimeOffset? TryGetProcessStartTimeUtc(int pid)
    {
        if (pid <= 0)
            return null;

        try
        {
            using var process = Process.GetProcessById(pid);
            if (process.HasExited)
                return null;

            return process.StartTime.ToUniversalTime();
        }
        catch (ArgumentException)
        {
            return null;
        }
        catch (InvalidOperationException)
        {
            return null;
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return null;
        }
    }
}
