using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;

namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// Atomically creates and cleans up the per-user ACL-protected connection descriptor file.
/// </summary>
public sealed class ConnectionDescriptorStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    private readonly string _descriptorPath;
    private readonly Func<SecurityIdentifier, FileSecurity> _fileSecurityFactory;

    public ConnectionDescriptorStore(
        string descriptorPath,
        Func<SecurityIdentifier, FileSecurity>? fileSecurityFactory = null)
    {
        if (string.IsNullOrWhiteSpace(descriptorPath))
            throw new ArgumentException("Descriptor path is required.", nameof(descriptorPath));

        _descriptorPath = descriptorPath;
        _fileSecurityFactory = fileSecurityFactory ?? BootstrapAcl.CreateFileSecurity;
    }

    public string DescriptorPath => _descriptorPath;

    /// <summary>
    /// Writes <paramref name="payload"/> atomically (temp file + replace) and applies
    /// a SID-restricted ACL. Never logs the token.
    /// </summary>
    public void WriteAtomic(SecureBootstrapPayload payload, SecurityIdentifier userSid)
    {
        ArgumentNullException.ThrowIfNull(payload);
        ArgumentNullException.ThrowIfNull(userSid);

        var directory = Path.GetDirectoryName(_descriptorPath)
            ?? throw new InvalidOperationException("Descriptor path has no directory.");

        Directory.CreateDirectory(directory);

        var json = JsonSerializer.Serialize(payload, JsonOptions);
        var bytes = Encoding.UTF8.GetBytes(json);
        var tempPath = _descriptorPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        var security = _fileSecurityFactory(userSid);

        try
        {
            // Write then ACL the temp file before publish so the token is never
            // world-readable as the final path.
            using (var stream = new FileStream(
                       tempPath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None,
                       bufferSize: 4096,
                       FileOptions.None))
            {
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush(flushToDisk: true);
            }

            FileSystemAclExtensions.SetAccessControl(new FileInfo(tempPath), security);
            File.Move(tempPath, _descriptorPath, overwrite: true);
            FileSystemAclExtensions.SetAccessControl(new FileInfo(_descriptorPath), security);
        }
        finally
        {
            TryDelete(tempPath);
        }
    }

    public SecureBootstrapPayload? TryRead()
    {
        if (!File.Exists(_descriptorPath))
            return null;

        try
        {
            var json = File.ReadAllText(_descriptorPath, Encoding.UTF8);
            return JsonSerializer.Deserialize<SecureBootstrapPayload>(json, JsonOptions);
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    /// <summary>
    /// Removes a stale descriptor when the owning process is dead. Returns true when removed.
    /// </summary>
    public bool TryCleanupStale(Func<int, bool>? isProcessAlive = null)
    {
        var existing = TryRead();
        if (existing is null)
        {
            // Corrupt / unreadable file still blocks restart — remove if present.
            if (File.Exists(_descriptorPath))
            {
                TryDelete(_descriptorPath);
                return true;
            }

            return false;
        }

        if (!StaleBootstrapDetector.IsStale(existing, isProcessAlive))
            return false;

        TryDelete(_descriptorPath);
        return true;
    }

    public void DeleteIfExists() => TryDelete(_descriptorPath);

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (IOException)
        {
            // Best-effort cleanup.
        }
        catch (UnauthorizedAccessException)
        {
            // Best-effort cleanup.
        }
    }
}
