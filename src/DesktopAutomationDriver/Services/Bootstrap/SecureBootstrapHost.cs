using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
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
        WriteIndented = false
    };

    private readonly SecureBootstrapPayload _payload;
    private readonly SecurityIdentifier _userSid;
    private readonly ConnectionDescriptorStore _descriptorStore;
    private readonly Func<SecurityIdentifier, PipeSecurity> _pipeSecurityFactory;
    private readonly ILogger _logger;
    private readonly CancellationTokenSource _cts = new();
    private Task? _listenTask;
    private int _started;
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
        _pipeSecurityFactory = pipeSecurityFactory ?? BootstrapAcl.CreatePipeSecurity;
        _logger = logger ?? NullLogger.Instance;

        _payload = new SecureBootstrapPayload
        {
            SchemaVersion = SecureBootstrapConstants.SchemaVersion,
            UserSid = sidValue,
            Username = driverContext.Username,
            Port = driverContext.MainPort,
            Token = driverContext.BearerToken,
            AuthorizationHeader = $"Bearer {driverContext.BearerToken}",
            Pid = Environment.ProcessId,
            StartedAtUtc = DateTimeOffset.UtcNow,
            PipeName = PipeName
        };
    }

    public string UserSid { get; }

    public string PipeName { get; }

    public string DescriptorPath => _descriptorStore.DescriptorPath;

    public bool IsRunning => Volatile.Read(ref _started) == 1;

    public Task StartAsync(CancellationToken cancellationToken = default)
    {
        if (Interlocked.CompareExchange(ref _started, 1, 0) != 0)
            throw new InvalidOperationException("Secure bootstrap host is already running.");

        // Remove crash leftovers before advertising a new descriptor / pipe.
        _descriptorStore.TryCleanupStale();

        // Fail fast if another live instance still owns the descriptor.
        var existing = _descriptorStore.TryRead();
        if (existing is not null && !StaleBootstrapDetector.IsStale(existing))
        {
            Interlocked.Exchange(ref _started, 0);
            throw new InvalidOperationException(
                $"A live secure bootstrap descriptor already exists for this user (pid {existing.Pid}). " +
                "Stop the other driver instance before starting a new one.");
        }

        _descriptorStore.WriteAtomic(_payload, _userSid);

        _logger.LogInformation(
            "Secure bootstrap active for SID {UserSid} via named pipe {PipeName} and descriptor {DescriptorPath}",
            UserSid,
            PipeName,
            DescriptorPath);

        _listenTask = Task.Run(() => ListenLoopAsync(_cts.Token), CancellationToken.None);
        return Task.CompletedTask;
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
                // Never include token material in error logs.
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
        // Idempotent: tests and await-using may dispose more than once.
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

        _descriptorStore.DeleteIfExists();

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
