namespace DesktopAutomationDriver.Services;

/// <summary>
/// Intentionally disabled. Same-user ACL file delivery is not human approval —
/// agents can self-read the file. Vision approval requires driver-owned interactive
/// presence (<see cref="Vision.IOperatorPresenceService"/>).
/// </summary>
[Obsolete("ACL operator-token files are not human presence. Use InteractiveOperatorPresenceService.")]
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
            "Operator ACL secret files are disabled. Vision approval requires interactive human presence.");

    /// <summary>Always returns null — ACL self-read is not an approval channel.</summary>
    public static string? TryRead(string? path = null) => null;
}
