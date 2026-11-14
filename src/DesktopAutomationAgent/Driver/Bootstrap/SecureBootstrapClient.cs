using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationAgent.Driver.Bootstrap;

/// <summary>
/// Discovers driver connection material via the SID ACL-protected named pipe
/// (preferred) or the ACL-protected connection descriptor file.
/// </summary>
public sealed class SecureBootstrapClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly Func<string> _currentSid;
    private readonly Func<string> _localAppData;
    private readonly Func<string, CancellationToken, Task<SecureBootstrapPayload?>>? _pipeReaderOverride;
    private readonly Func<string, SecureBootstrapPayload?>? _descriptorReaderOverride;
    private readonly Func<int, DateTimeOffset?>? _getProcessStartTimeUtc;
    private readonly ILogger _logger;
    private readonly TimeSpan _pipeConnectTimeout;

    public SecureBootstrapClient(
        TimeSpan? pipeConnectTimeout = null,
        ILogger? logger = null)
        : this(
            WindowsUserSid.RequireCurrent,
            static () => Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            pipeReaderOverride: null,
            descriptorReaderOverride: null,
            getProcessStartTimeUtc: null,
            pipeConnectTimeout,
            logger)
    {
    }

    internal SecureBootstrapClient(
        Func<string> currentSid,
        Func<string> localAppData,
        Func<string, CancellationToken, Task<SecureBootstrapPayload?>>? pipeReaderOverride,
        Func<string, SecureBootstrapPayload?>? descriptorReaderOverride,
        Func<int, DateTimeOffset?>? getProcessStartTimeUtc,
        TimeSpan? pipeConnectTimeout = null,
        ILogger? logger = null)
    {
        _currentSid = currentSid;
        _localAppData = localAppData;
        _pipeReaderOverride = pipeReaderOverride;
        _descriptorReaderOverride = descriptorReaderOverride;
        _getProcessStartTimeUtc = getProcessStartTimeUtc;
        _pipeConnectTimeout = pipeConnectTimeout ?? TimeSpan.FromSeconds(3);
        _logger = logger ?? NullLogger.Instance;
    }

    public async Task<SecureBootstrapPayload> DiscoverAsync(CancellationToken cancellationToken = default)
    {
        if (!OperatingSystem.IsWindows() && _pipeReaderOverride is null && _descriptorReaderOverride is null)
        {
            throw new DriverConnectionException(
                "Secure driver discovery requires Windows (SID ACL named pipe / connection descriptor).");
        }

        var sid = _currentSid();
        if (string.IsNullOrWhiteSpace(sid))
            throw new DriverConnectionException("Current Windows user SID is unavailable for secure discovery.");

        var pipeName = SecureBootstrapConstants.BuildPipeName(sid);
        _logger.LogInformation(
            "Discovering driver via secure bootstrap for SID {UserSid}",
            sid);

        SecureBootstrapPayload? payload = null;
        Exception? pipeError = null;

        try
        {
            payload = _pipeReaderOverride is not null
                ? await _pipeReaderOverride(pipeName, cancellationToken).ConfigureAwait(false)
                : await ReadFromPipeAsync(pipeName, sid, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested)
        {
            // Connect/read timeout from our own CancelAfter — not caller cancellation.
            pipeError = ex;
            _logger.LogDebug(ex, "Secure bootstrap named pipe timed out; trying descriptor file.");
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not DriverConnectionException)
        {
            pipeError = ex;
            _logger.LogDebug(ex, "Secure bootstrap named pipe was unavailable; trying descriptor file.");
        }

        if (payload is null)
        {
            var descriptorPath = SecureBootstrapConstants.GetDefaultDescriptorPath(_localAppData());
            payload = _descriptorReaderOverride is not null
                ? _descriptorReaderOverride(descriptorPath)
                : TryReadDescriptor(descriptorPath);

            if (payload is null)
            {
                var detail = pipeError is null
                    ? "named pipe and connection descriptor were unavailable"
                    : $"named pipe failed ({pipeError.GetType().Name}) and connection descriptor was unavailable";

                throw new DriverConnectionException(
                    $"Unable to discover the desktop automation driver via secure bootstrap ({detail}). " +
                    "Is the driver running under the same Windows user?");
            }
        }

        ValidatePayload(payload, sid, authenticatedServerProcessId: null);
        return payload;
    }

    private async Task<SecureBootstrapPayload?> ReadFromPipeAsync(
        string pipeName,
        string expectedSid,
        CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        linked.CancelAfter(_pipeConnectTimeout);

        await using var client = new NamedPipeClientStream(
            ".",
            pipeName,
            PipeDirection.In,
            PipeOptions.Asynchronous);

        await client.ConnectAsync(linked.Token).ConfigureAwait(false);

        // Authenticate the remote server process BEFORE reading/trusting any payload bytes.
        if (!OperatingSystem.IsWindows())
        {
            throw new DriverConnectionException(
                "Secure bootstrap named-pipe authentication requires Windows.");
        }

        if (!NamedPipeServerAuthenticator.TryGetServerIdentity(client, out var serverPid, out var serverSid)
            || string.IsNullOrWhiteSpace(serverSid))
        {
            throw new DriverConnectionException(
                "Unable to authenticate the secure bootstrap named-pipe server process.");
        }

        if (!string.Equals(serverSid.Trim(), expectedSid.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            throw new DriverConnectionException(
                "Secure bootstrap named-pipe server process SID does not match the current Windows user SID.");
        }

        using var reader = new StreamReader(client, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
        var json = await reader.ReadToEndAsync(linked.Token).ConfigureAwait(false);
        if (string.IsNullOrWhiteSpace(json))
            return null;

        var payload = JsonSerializer.Deserialize<SecureBootstrapPayload>(json, JsonOptions);
        if (payload is null)
            return null;

        // Re-validate claims against the OS-authenticated server identity.
        ValidatePayload(payload, expectedSid, authenticatedServerProcessId: serverPid);
        return payload;
    }

    private static SecureBootstrapPayload? TryReadDescriptor(string path)
    {
        if (!File.Exists(path))
            return null;

        try
        {
            var json = File.ReadAllText(path, Encoding.UTF8);
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

    private void ValidatePayload(
        SecureBootstrapPayload payload,
        string expectedSid,
        int? authenticatedServerProcessId)
    {
        if (StaleBootstrapDetector.IsStale(payload, _getProcessStartTimeUtc))
        {
            throw new DriverConnectionException(
                "Secure bootstrap descriptor is stale (owning driver process is not running or PID was reused). " +
                "Start the desktop automation driver and retry.");
        }

        if (!string.Equals(payload.UserSid.Trim(), expectedSid.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            throw new DriverConnectionException(
                "Secure bootstrap payload SID does not match the current Windows user SID.");
        }

        if (authenticatedServerProcessId is int serverPid && payload.Pid != serverPid)
        {
            throw new DriverConnectionException(
                "Secure bootstrap payload PID does not match the authenticated named-pipe server process.");
        }

        if (payload.Port <= 0)
            throw new DriverConnectionException("Secure bootstrap payload returned an invalid port.");

        if (string.IsNullOrWhiteSpace(payload.Token) && string.IsNullOrWhiteSpace(payload.AuthorizationHeader))
            throw new DriverConnectionException("Secure bootstrap payload did not include a token.");
    }
}
