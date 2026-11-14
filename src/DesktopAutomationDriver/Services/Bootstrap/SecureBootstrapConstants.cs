namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>Shared naming for the SID-scoped secure bootstrap channel.</summary>
public static class SecureBootstrapConstants
{
    public const int SchemaVersion = 1;

    /// <summary>Named-pipe name prefix; full name is "{prefix}.{userSid}".</summary>
    public const string PipeNamePrefix = "DesktopAutomationDriver.bootstrap";

    public const string RelativeDirectoryName = "DesktopAutomationDriver";
    public const string BootstrapFolderName = "bootstrap";
    public const string DescriptorFileName = "connection.json";
    public const string HostLockFileName = "host.lock";

    /// <summary>
    /// Allowed skew when comparing recorded process start time to a live process
    /// (Windows StartTime resolution / UTC conversion).
    /// </summary>
    public static readonly TimeSpan ProcessStartTimeTolerance = TimeSpan.FromSeconds(2);

    public static string BuildPipeName(string userSid)
    {
        if (string.IsNullOrWhiteSpace(userSid))
            throw new ArgumentException("User SID is required.", nameof(userSid));

        // Pipe names cannot contain backslash; SIDs use hyphens and are safe as-is.
        return $"{PipeNamePrefix}.{userSid.Trim()}";
    }

    public static string GetDefaultDescriptorDirectory(string localAppDataRoot)
    {
        if (string.IsNullOrWhiteSpace(localAppDataRoot))
            throw new ArgumentException("LOCALAPPDATA root is required.", nameof(localAppDataRoot));

        return Path.Combine(localAppDataRoot, RelativeDirectoryName, BootstrapFolderName);
    }

    public static string GetDefaultDescriptorPath(string localAppDataRoot) =>
        Path.Combine(GetDefaultDescriptorDirectory(localAppDataRoot), DescriptorFileName);
}
