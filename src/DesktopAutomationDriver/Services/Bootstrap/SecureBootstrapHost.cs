using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationDriver.Services.Bootstrap;

/// <summary>
/// Serves secure bootstrap payloads over a SID-restricted named pipe and
/// maintains an ACL-protected on-disk connection descriptor for crash recovery.
/// </summary>
public sealed class SecureBootstrapHost : ISecureBootstrapHost
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
        // Descriptor/pipe payloads are not HTML; keep base64 tokens unescaped (+,/).
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private readonly SecureBootstrapPayload _payload;
    private readonly SecurityIdentifier _userSid;
    private readonly ConnectionDescriptorStore _descriptorStore;
    private readonly string _lockPath;
    private readonly Func<SecurityIdentifier, PipeSecurity> _pipeSecurityFactory;
    private readonly Func<SecurityIdentifier, FileSecurity> _fileSecurityFactory;
    private readonly ILogger _logger;
    private readonly CancellationTokenSource _cts = new();
    private FileStream? _exclusiveLock;
    private Task? _listenTask;
    private int _started;
    private int _published;
    private int _disposed;

    public SecureBootstrapHost(
        IDriverContext driverContext,
        string? descriptorPath = null,
        Func<SecurityIdentifier, PipeSecurity>? pipeSecurityFactory = null,
        Func<SecurityIdentifier, FileSecurity>? fileSecurityFactory = null,
        ILogger? logger = null)
    {
        ArgumentNullException.ThrowIfNull(driverContext);

        var sidValue = string.IsNullOrWhiteSpace(driverContext.UserSid)
            ? WindowsUserSid.RequireCurrent()
            : driverContext.UserSid;

        UserSid = sidValue;
        _userSid = WindowsUserSid.Parse(sidValue);
        PipeName = SecureBootstrapConstants.BuildPipeName(sidValue);

        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var path = descriptorPath
            ?? SecureBootstrapConstants.GetDefaultDescriptorPath(localAppData);

        _descriptorStore = new ConnectionDescriptorStore(path, fileSecurityFactory);
        _lockPath = Path.Combine(
            Path.GetDirectoryName(path)
            ?? throw new InvalidOperationException("Descriptor path has no directory."),
            SecureBootstrapConstants.HostLockFileName);
        _pipeSecurityFactory = pipeSecurityFactory ?? BootstrapAcl.CreatePipeSecurity;
        _fileSecurityFactory = fileSecurityFactory ?? BootstrapAcl.CreateFileSecurity;
        _logger = logger ?? NullLogger.Instance;

        var processStart = StaleBootstrapDetector.GetCurrentProcessStartTimeUtc();
        _payload = new SecureBootstrapPayload
        {
            SchemaVersion = SecureBootstrapConstants.SchemaVersion,
            UserSid = sidValue,
            Username = driverContext.Username,
            Port = driverContext.MainPort,
            Token = driverContext.BearerToken,
            AuthorizationHeader = $"Bearer {driverContext.BearerToken}",
            Pid = Environment.ProcessId,
            ProcessStartTimeUtc = processStart,
            StartedAtUtc = DateTimeOffset.UtcNow,
            PipeName = PipeName
        };
    }

    public string UserSid { get; }

    public string PipeName { get; }

    public string DescriptorPath => _descriptorStore.DescriptorPath;

    public bool IsRunning => Volatile.Read(ref _started) == 1;

    /// <summary>True when this host successfully published the live descriptor.</summary>
    public bool HasPublishedDescriptor => Volatile.Read(ref _published) == 1;

    /// <summary>PID recorded in the published payload (0 if never published).</summary>
    public int PublishedPid => HasPublishedDescriptor ? _payload.Pid : 0;

    /// <summary>Process start time recorded in the published payload.</summary>
    public DateTimeOffset PublishedProcessStartTimeUtc =>
        HasPublishedDescriptor ? _payload.ProcessStartTimeUtc : default;

    public Task StartAsync(CancellationToken cancellationToken = default)
    {
        if (Interlocked.CompareExchange(ref _started, 1, 0) != 0)
            throw new InvalidOperationException("Secure bootstrap host is already running.");

        // Rejected hosts must never enter the publish/rollback path: no WriteAtomic,
        // no DeleteIfExists, and no catch that could touch a live peer descriptor.
        if (!TryAcquireExclusiveLock())
        {
            Interlocked.Exchange(ref _started, 0);
            throw new InvalidOperationException(
                "A live secure bootstrap instance already holds the per-user lock for this SID. " +
                "Stop the other driver instance before starting a new one.");
        }

        try
        {
            // Remove crash leftovers only after we own the lock.
            _descriptorStore.TryCleanupStale();

            var existing = _descriptorStore.TryRead();
            if (existing is not null && !StaleBootstrapDetector.IsStale(existing))
            {
                // Live peer descriptor — release lock only; never delete/overwrite it.
                ReleaseExclusiveLock();
                Interlocked.Exchange(ref _started, 0);
                throw new InvalidOperationException(
                    $"A live secure bootstrap descriptor already exists for this user (pid {existing.Pid}). " +
                    "Stop the other driver instance before starting a new one.");
            }

            _descriptorStore.WriteAtomic(_payload, _userSid);
            Interlocked.Exchange(ref _published, 1);

            _logger.LogInformation(
                "Secure bootstrap active for SID {UserSid} via named pipe {PipeName} and descriptor {DescriptorPath}",
                UserSid,
                PipeName,
                DescriptorPath);

            _listenTask = Task.Run(() => ListenLoopAsync(_cts.Token), CancellationToken.None);
            return Task.CompletedTask;
        }
        catch
        {
            // Roll back only a descriptor this host just published. Never delete a
            // peer's live file (rejected hosts never set _published).
            if (Interlocked.Exchange(ref _published, 0) == 1)
                _descriptorStore.DeleteIfExists();

            ReleaseExclusiveLock();
            Interlocked.Exchange(ref _started, 0);
            throw;
        }
    }

    private bool TryAcquireExclusiveLock()
    {
        var directory = Path.GetDirectoryName(_lockPath)
            ?? throw new InvalidOperationException("Lock path has no directory.");
        Directory.CreateDirectory(directory);

        try
        {
            var security = _fileSecurityFactory(_userSid);
            var info = new FileInfo(_lockPath);
            // Create or open with exclusive share — only one live host can hold it.
            _exclusiveLock = FileSystemAclExtensions.Create(
                info,
                FileMode.OpenOrCreate,
                FileSystemRights.Read | FileSystemRights.Write,
                FileShare.None,
                bufferSize: 1,
                FileOptions.DeleteOnClose,
                security);
            return true;
        }
        catch (IOException)
        {
            _exclusiveLock = null;
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            _exclusiveLock = null;
            return false;
        }
    }

    private void ReleaseExclusiveLock()
    {
        var lockStream = Interlocked.Exchange(ref _exclusiveLock, null);
        if (lockStream is null)
            return;

        try { lockStream.Dispose(); } catch { /* best effort */ }
    }

    private async Task ListenLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            NamedPipeServerStream? server = null;
            try
            {
                server = CreateServerStream();
                await server.WaitForConnectionAsync(cancellationToken).ConfigureAwait(false);

                var json = JsonSerializer.Serialize(_payload, JsonOptions);
                var bytes = Encoding.UTF8.GetBytes(json);
                await server.WriteAsync(bytes, cancellationToken).ConfigureAwait(false);
                await server.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                break;
            }
            catch (IOException ex)
            {
                _logger.LogDebug(ex, "Secure bootstrap pipe connection ended.");
            }
            catch (UnauthorizedAccessException ex)
            {
                _logger.LogWarning(ex, "Secure bootstrap pipe access denied for SID {UserSid}.", UserSid);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Secure bootstrap pipe listener error for SID {UserSid}.", UserSid);
                await Task.Delay(50, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                if (server is not null)
                {
                    try { server.Dispose(); } catch { /* best effort */ }
                }
            }
        }
    }

    private NamedPipeServerStream CreateServerStream()
    {
        var security = _pipeSecurityFactory(_userSid);
        return NamedPipeServerStreamAcl.Create(
            PipeName,
            PipeDirection.Out,
            maxNumberOfServerInstances: 1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            inBufferSize: 0,
            outBufferSize: 64 * 1024,
            pipeSecurity: security);
    }

    public async ValueTask DisposeAsync()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0)
            return;

        try
        {
            _cts.Cancel();
        }
        catch (ObjectDisposedException)
        {
            // Already canceled/disposed by a racing shutdown path.
        }

        if (_listenTask is not null)
        {
            try { await _listenTask.ConfigureAwait(false); }
            catch { /* ignore shutdown races */ }
        }

        // Only the host that successfully published may delete the live descriptor.
        // A rejected second StartAsync must not wipe the first instance's file.
        if (Interlocked.Exchange(ref _published, 0) == 1)
            _descriptorStore.DeleteIfExists();

        ReleaseExclusiveLock();

        try
        {
            _cts.Dispose();
        }
        catch (ObjectDisposedException)
        {
            // Best-effort.
        }

        Interlocked.Exchange(ref _started, 0);
        GC.SuppressFinalize(this);
    }
}
