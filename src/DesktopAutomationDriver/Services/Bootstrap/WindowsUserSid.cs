using System.Security.Principal;

namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>Resolves the Windows SID for the current process identity.</summary>
public static class WindowsUserSid
{
    /// <summary>
    /// Returns the current Windows user SID string (e.g. S-1-5-21-...), or null
    /// when the platform cannot provide one.
    /// </summary>
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
                "Unable to resolve the current Windows user SID for secure bootstrap.");
        }

        return sid;
    }

    public static SecurityIdentifier Parse(string sid)
    {
        if (string.IsNullOrWhiteSpace(sid))
            throw new ArgumentException("SID is required.", nameof(sid));

        return new SecurityIdentifier(sid.Trim());
    }
}
