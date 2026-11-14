namespace DesktopAutomationAgent.Driver.Bootstrap;

/// <summary>Shared naming for the SID-scoped secure bootstrap channel (must match the driver).</summary>
public static class SecureBootstrapConstants
{
    public const int SchemaVersion = 1;

    public const string PipeNamePrefix = "DesktopAutomationDriver.bootstrap";

    public const string RelativeDirectoryName = "DesktopAutomationDriver";
    public const string BootstrapFolderName = "bootstrap";
    public const string DescriptorFileName = "connection.json";
    public const string HostLockFileName = "host.lock";

    public static string BuildPipeName(string userSid)
    {
        if (string.IsNullOrWhiteSpace(userSid))
            throw new ArgumentException("User SID is required.", nameof(userSid));

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
