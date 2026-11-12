using System.Net;
using System.Net.Http.Headers;
using DesktopAutomationAgent.Cli;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver;
using Microsoft.Extensions.Logging;

namespace DesktopAutomationAgent.Jira;

/// <summary>
/// Mockable Jira HTTP transport: bounded timeouts (headers + body), response-size limits,
/// idempotent GET retries, rate-limit reporting. Never logs Authorization.
/// </summary>
public static class JiraHttpClientNames
{
    public const string Default = "jira";
}

internal sealed class JiraHttpTransport
{
    public const string HttpClientName = JiraHttpClientNames.Default;

    private readonly IHttpClientFactory _httpClientFactory;
    private readonly JiraOptions _options;
    private readonly ILogger _logger;
    private readonly AuthenticationHeaderValue _auth;
    private readonly Uri _baseUri;

    public JiraHttpTransport(
        IHttpClientFactory httpClientFactory,
        JiraOptions options,
        ILogger logger)
    {
        _httpClientFactory = httpClientFactory;
        _options = options;
        _logger = logger;
        _baseUri = JiraUrlRules.EnsureAllowed(options.BaseUrl!, options.AllowLocalHttpTestServer);
        _auth = JiraAuthResolver.Resolve(options);
    }

    public async Task<JiraHttpResponse> GetAsync(
        string relativeOrAbsolutePath,
        IDictionary<string, string>? conditionalHeaders,
        CancellationToken cancellationToken)
    {
        var attempt = 0;
        JiraHttpResponse? last = null;
        // MaxRetries = additional retries after the first attempt.
        // Total attempts = MaxRetries + 1. Stop when attempt >= MaxRetries + 1.
        var maxAttempts = _options.MaxRetries + 1;
        while (true)
        {
            attempt++;
            last = await SendOnceAsync(relativeOrAbsolutePath, conditionalHeaders, cancellationToken)
                .ConfigureAwait(false);

            if (!ShouldRetry(last) || attempt >= maxAttempts)
                return last;

            var delay = ComputeRetryDelay(last, attempt);
            _logger.LogWarning(
                "Jira GET retry {Attempt}/{MaxAttempts} after status {Status} (delay {DelayMs}ms). Authorization redacted.",
                attempt,
                maxAttempts,
                (int)last.StatusCode,
                (int)delay.TotalMilliseconds);

            await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task<JiraHttpResponse> SendOnceAsync(
        string relativeOrAbsolutePath,
        IDictionary<string, string>? conditionalHeaders,
        CancellationToken cancellationToken)
    {
        using var client = _httpClientFactory.CreateClient(HttpClientName);
        client.BaseAddress = _baseUri;
        // Cover headers + body with one linked timeout; HttpClient.Timeout alone
        // does not bound ResponseHeadersRead body streaming.
        client.Timeout = Timeout.InfiniteTimeSpan;

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(TimeSpan.FromSeconds(_options.RequestTimeoutSeconds));
        var timedToken = timeoutCts.Token;

        using var request = new HttpRequestMessage(HttpMethod.Get, relativeOrAbsolutePath);
        request.Headers.Authorization = _auth;
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (conditionalHeaders is not null)
        {
            foreach (var pair in conditionalHeaders)
                request.Headers.TryAddWithoutValidation(pair.Key, pair.Value);
        }

        HttpResponseMessage response;
        try
        {
            response = await client.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead,
                    timedToken)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            return new JiraHttpResponse
            {
                StatusCode = 0,
                Errors = ["Jira request timed out."],
                ExitCode = ExitCodes.ExecutionFailure
            };
        }
        catch (HttpRequestException ex)
        {
            return new JiraHttpResponse
            {
                StatusCode = 0,
                Errors = [$"Jira transport failure: {SecretRedactor.Redact(ex.Message)}"],
                ExitCode = ExitCodes.DriverUnavailable
            };
        }

        using (response)
        {
            var rate = ReadRateLimit(response);
            if (response.StatusCode == HttpStatusCode.NotModified)
            {
                return new JiraHttpResponse
                {
                    StatusCode = response.StatusCode,
                    NotModified = true,
                    ETag = response.Headers.ETag?.Tag,
                    LastModified = response.Content.Headers.LastModified?.ToString("R"),
                    RateLimit = rate,
                    ExitCode = ExitCodes.Success
                };
            }

            byte[]? bytes;
            try
            {
                bytes = await ReadBoundedAsync(response.Content, timedToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                return new JiraHttpResponse
                {
                    StatusCode = response.StatusCode,
                    Errors = ["Jira response body read timed out."],
                    RateLimit = rate,
                    ExitCode = ExitCodes.ExecutionFailure
                };
            }

            if (bytes is null)
            {
                return new JiraHttpResponse
                {
                    StatusCode = response.StatusCode,
                    Errors = [$"Jira response exceeded MaxResponseBytes ({_options.MaxResponseBytes})."],
                    RateLimit = rate,
                    ExitCode = ExitCodes.SuiteOrWorkspace
                };
            }

            var exit = MapStatus(response.StatusCode);
            var errors = exit == ExitCodes.Success
                ? Array.Empty<string>()
                : new[] { $"Jira HTTP {(int)response.StatusCode} ({response.ReasonPhrase})." };

            return new JiraHttpResponse
            {
                StatusCode = response.StatusCode,
                Body = bytes,
                ETag = response.Headers.ETag?.Tag,
                LastModified = response.Content.Headers.LastModified?.ToString("R"),
                RateLimit = rate,
                Errors = errors,
                ExitCode = exit
            };
        }
    }

    private async Task<byte[]?> ReadBoundedAsync(HttpContent content, CancellationToken cancellationToken)
    {
        await using var stream = await content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        using var ms = new MemoryStream();
        var buffer = new byte[8192];
        long total = 0;
        while (true)
        {
            var read = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken)
                .ConfigureAwait(false);
            if (read == 0)
                break;
            total += read;
            if (total > _options.MaxResponseBytes)
                return null;
            ms.Write(buffer, 0, read);
        }

        return ms.ToArray();
    }

    private static bool ShouldRetry(JiraHttpResponse response) =>
        response.StatusCode is HttpStatusCode.RequestTimeout
            or HttpStatusCode.TooManyRequests
            or HttpStatusCode.BadGateway
            or HttpStatusCode.ServiceUnavailable
            or HttpStatusCode.GatewayTimeout
        || response.StatusCode == 0;

    private static TimeSpan ComputeRetryDelay(JiraHttpResponse response, int attempt)
    {
        if (response.RateLimit?.RetryAfterSeconds is int retryAfter && retryAfter > 0)
            return TimeSpan.FromSeconds(Math.Min(retryAfter, 30));

        // Keep unit tests fast: small bounded backoff.
        return TimeSpan.FromMilliseconds(Math.Min(50 * attempt * attempt, 200));
    }

    private static JiraRateLimitInfo ReadRateLimit(HttpResponseMessage response)
    {
        int? limit = TryHeaderInt(response, "X-RateLimit-Limit");
        int? remaining = TryHeaderInt(response, "X-RateLimit-Remaining");
        int? retryAfter = null;
        if (response.Headers.RetryAfter?.Delta is { } delta)
            retryAfter = (int)Math.Ceiling(delta.TotalSeconds);
        else if (TryHeaderInt(response, "Retry-After") is int seconds)
            retryAfter = seconds;

        return new JiraRateLimitInfo
        {
            Limit = limit,
            Remaining = remaining,
            RetryAfterSeconds = retryAfter
        };
    }

    private static int? TryHeaderInt(HttpResponseMessage response, string name)
    {
        if (!response.Headers.TryGetValues(name, out var values)
            && !response.Content.Headers.TryGetValues(name, out values))
            return null;

        var raw = values.FirstOrDefault();
        return int.TryParse(raw, out var n) ? n : null;
    }

    private static int MapStatus(HttpStatusCode status) =>
        status switch
        {
            HttpStatusCode.OK => ExitCodes.Success,
            HttpStatusCode.Unauthorized => ExitCodes.AuthOrCatalog,
            HttpStatusCode.Forbidden => ExitCodes.AuthOrCatalog,
            HttpStatusCode.NotFound => ExitCodes.SuiteOrWorkspace,
            HttpStatusCode.TooManyRequests => ExitCodes.ExecutionFailure,
            _ when (int)status >= 500 => ExitCodes.DriverUnavailable,
            _ => ExitCodes.SuiteOrWorkspace
        };
}

internal sealed class JiraHttpResponse
{
    public HttpStatusCode StatusCode { get; init; }

    public byte[]? Body { get; init; }

    public bool NotModified { get; init; }

    public string? ETag { get; init; }

    public string? LastModified { get; init; }

    public JiraRateLimitInfo? RateLimit { get; init; }

    public IReadOnlyList<string> Errors { get; init; } = [];

    public int ExitCode { get; init; }
}
