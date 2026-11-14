using System.Security.Principal;

namespace DesktopAutomationAgent.Driver.Bootstrap;

public static class WindowsUserSid
{
    public static string? TryGetCurrent()
    {
        if (!OperatingSystem.IsWindows())
            return null;

        try
        {
            using var identity = WindowsIdentity.GetCurrent();
            return identity.User?.Value;
        }
        catch
        {
            return null;
        }
    }

    public static string RequireCurrent()
    {
        var sid = TryGetCurrent();
        if (string.IsNullOrWhiteSpace(sid))
        {
            throw new InvalidOperationException(
                "Unable to resolve the current Windows user SID for secure driver discovery.");
        }

        return sid;
    }
}
