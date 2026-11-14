namespace DesktopAutomationDriver.Services;

/// <summary>
/// Intentionally disabled. Same-user ACL file delivery is not an approval channel.
/// Vision approval uses operator confirmation UX
/// (<see cref="Vision.IOperatorPresenceService"/>) — UIA-automatable, not OS presence.
/// </summary>
[Obsolete("ACL operator-token files are disabled. Use InteractiveOperatorPresenceService (UX confirmation).")]
public static class OperatorTokenAclSecret
{
    public const string RelativeDirectoryName = "DesktopAutomationDriver";
    public const string FileName = "operator-token";

    public static string GetDefaultPath()
    {
        var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(root))
            throw new InvalidOperationException("LOCALAPPDATA is unavailable.");

        return Path.Combine(root, RelativeDirectoryName, FileName);
    }

    /// <summary>Always throws — ACL self-read delivery is removed.</summary>
    public static string WriteCurrentUserOnly(string token, string? path = null) =>
        throw new InvalidOperationException(
            "Operator ACL secret files are disabled. Vision approval uses operator confirmation UX.");

    /// <summary>Always returns null — ACL self-read is not an approval channel.</summary>
    public static string? TryRead(string? path = null) => null;
}
