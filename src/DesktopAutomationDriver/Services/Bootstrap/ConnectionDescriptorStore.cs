using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Encodings.Web;
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
        WriteIndented = false,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
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
    /// Writes <paramref name="payload"/> atomically.
    /// Temp file is created with restrictive <see cref="FileSecurity"/> before any
    /// token bytes are written. <see cref="File.Move"/> onto the final path is the
    /// publication point; any failure after rename rolls back by deleting the final
    /// path (caller must hold the per-user host lock).
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
        var published = false;

        try
        {
            var tempInfo = new FileInfo(tempPath);
            using (var stream = FileSystemAclExtensions.Create(
                       tempInfo,
                       FileMode.CreateNew,
                       FileSystemRights.Write | FileSystemRights.Read,
                       FileShare.None,
                       bufferSize: 4096,
                       FileOptions.None,
                       security))
            {
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush(flushToDisk: true);
            }

            // Rename is the final publication — NTFS preserves the temp ACL.
            File.Move(tempPath, _descriptorPath, overwrite: true);
            published = true;

            try
            {
                // Defense-in-depth: re-assert ACL. On failure, roll back the published file.
                FileSystemAclExtensions.SetAccessControl(new FileInfo(_descriptorPath), security);
            }
            catch
            {
                TryDelete(_descriptorPath);
                published = false;
                throw;
            }
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
    /// Removes a stale descriptor when the owning process is dead / PID-reused.
    /// </summary>
    public bool TryCleanupStale(Func<int, DateTimeOffset?>? getProcessStartTimeUtc = null)
    {
        var existing = TryRead();
        if (existing is null)
        {
            if (File.Exists(_descriptorPath))
            {
                TryDelete(_descriptorPath);
                return true;
            }

            return false;
        }

        if (!StaleBootstrapDetector.IsStale(existing, getProcessStartTimeUtc))
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
