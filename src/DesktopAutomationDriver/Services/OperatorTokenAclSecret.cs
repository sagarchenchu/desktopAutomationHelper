using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Delivers the vision operator token via a current-user ACL-protected local file.
/// The token must never be written to stdout or standard logger output.
/// </summary>
public static class OperatorTokenAclSecret
{
    public const string RelativeDirectoryName = "DesktopAutomationDriver";
    public const string FileName = "operator-token";

    /// <summary>
    /// Default path: %LOCALAPPDATA%\DesktopAutomationDriver\operator-token
    /// </summary>
    public static string GetDefaultPath()
    {
        var root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(root))
            throw new InvalidOperationException("LOCALAPPDATA is unavailable for operator token ACL secret.");

        return Path.Combine(root, RelativeDirectoryName, FileName);
    }

    /// <summary>
    /// Atomically writes <paramref name="token"/> to an ACL-protected file readable only by
    /// the current Windows user (and Local System). Returns the final path.
    /// </summary>
    public static string WriteCurrentUserOnly(string token, string? path = null)
    {
        if (string.IsNullOrWhiteSpace(token))
            throw new ArgumentException("Operator token must be non-empty.", nameof(token));

        path ??= GetDefaultPath();
        var directory = Path.GetDirectoryName(path)
            ?? throw new InvalidOperationException("Operator token ACL path has no directory.");

        Directory.CreateDirectory(directory);
        HardenDirectoryAcl(directory);

        var tempPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllText(tempPath, token.Trim(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            ApplyCurrentUserOnlyAcl(tempPath);
            File.Move(tempPath, path, overwrite: true);
            ApplyCurrentUserOnlyAcl(path);
            return path;
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                try { File.Delete(tempPath); } catch { /* best effort */ }
            }
        }
    }

    /// <summary>Reads the ACL secret if present (for operator UI / acceptance harness).</summary>
    public static string? TryRead(string? path = null)
    {
        path ??= GetDefaultPath();
        if (!File.Exists(path))
            return null;

        var text = File.ReadAllText(path, Encoding.UTF8).Trim();
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private static void HardenDirectoryAcl(string directory)
    {
        var dirInfo = new DirectoryInfo(directory);
        var security = new DirectorySecurity();
        security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
        foreach (var rule in CreateCurrentUserAllowRules())
            security.AddAccessRule(rule);
        dirInfo.SetAccessControl(security);
    }

    private static void ApplyCurrentUserOnlyAcl(string filePath)
    {
        var fileInfo = new FileInfo(filePath);
        var security = new FileSecurity();
        security.SetAccessRuleProtection(isProtected: true, preserveInheritance: false);
        foreach (var rule in CreateCurrentUserAllowRules())
            security.AddAccessRule(rule);
        fileInfo.SetAccessControl(security);
    }

    private static IEnumerable<FileSystemAccessRule> CreateCurrentUserAllowRules()
    {
        var user = WindowsIdentity.GetCurrent().User
            ?? throw new InvalidOperationException("Current Windows user SID is unavailable.");

        yield return new FileSystemAccessRule(
            user,
            FileSystemRights.FullControl,
            InheritanceFlags.None,
            PropagationFlags.None,
            AccessControlType.Allow);

        yield return new FileSystemAccessRule(
            new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null),
            FileSystemRights.FullControl,
            InheritanceFlags.None,
            PropagationFlags.None,
            AccessControlType.Allow);
    }
}
