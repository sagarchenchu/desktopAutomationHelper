using System.Diagnostics;

namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// Detects stale bootstrap descriptors left behind after a driver crash.
/// Stale = missing process, or PID alive but OS process creation time does not
/// exactly match the recorded value (PID reuse).
/// </summary>
public static class StaleBootstrapDetector
{
    /// <summary>
    /// Returns true when the descriptor should be ignored / cleaned up.
    /// </summary>
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

        if (string.IsNullOrWhiteSpace(payload.Token))
            return true;

        var lookup = getProcessStartTimeUtc ?? TryGetProcessStartTimeUtc;
        var actual = lookup(payload.Pid);
        if (actual is null)
            return true;

        // Exact OS creation-time match — no skew tolerance.
        return actual.Value.UtcTicks != payload.ProcessStartTimeUtc.UtcTicks;
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

            return ToUtcOffset(process.StartTime);
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

    public static DateTimeOffset GetCurrentProcessStartTimeUtc()
    {
        using var process = Process.GetCurrentProcess();
        return ToUtcOffset(process.StartTime);
    }

    private static DateTimeOffset ToUtcOffset(DateTime startTime)
    {
        var utc = startTime.Kind == DateTimeKind.Utc
            ? startTime
            : startTime.ToUniversalTime();
        return new DateTimeOffset(utc.Ticks, TimeSpan.Zero);
    }
}
