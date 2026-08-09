using System.Net.Http.Json;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Driver;

public sealed class DriverConnectionResolver : IDriverConnectionResolver
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly DriverOptions _options;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<DriverConnectionResolver> _logger;
    private readonly Func<string> _currentUserName;

    public DriverConnectionResolver(
        IOptions<AgentOptions> options,
        IHttpClientFactory httpClientFactory,
        ILogger<DriverConnectionResolver> logger)
        : this(options, httpClientFactory, logger, static () => Environment.UserName)
    {
    }

    internal DriverConnectionResolver(
        IOptions<AgentOptions> options,
        IHttpClientFactory httpClientFactory,
        ILogger<DriverConnectionResolver> logger,
        Func<string> currentUserName)
    {
        _options = options.Value.Driver;
        _httpClientFactory = httpClientFactory;
        _logger = logger;
        _currentUserName = currentUserName;
    }

    public async Task<DriverConnection> ResolveAsync(CancellationToken cancellationToken = default)
    {
        AgentOptionsValidator.ValidateDriverOptions(_options);

        var hasUrl = !string.IsNullOrWhiteSpace(_options.BaseUrl);
        var hasToken = !string.IsNullOrWhiteSpace(_options.BearerToken);

        if (hasUrl && hasToken)
        {
            var explicitUri = DriverUrlRules.EnsureAllowed(_options.BaseUrl!, _options.AllowRemoteDriver);
            _logger.LogInformation(
                "Using explicit driver connection at {BaseUrl}",
                $"{explicitUri.Scheme}://{explicitUri.Authority}");

            return new DriverConnection
            {
                BaseUri = NormalizeBaseUri(explicitUri),
                BearerToken = _options.BearerToken!,
                DiscoveryMethod = "explicit"
            };
        }

        return await DiscoverViaVerifyAsync(cancellationToken).ConfigureAwait(false);
    }

    private async Task<DriverConnection> DiscoverViaVerifyAsync(CancellationToken cancellationToken)
    {
        // Discovery is loopback-only. Remote drivers must be configured explicitly.
        var verifyUri = DriverUrlRules.EnsureAllowed(_options.VerifyUrl, allowRemote: false);
        if (!DriverUrlRules.IsLoopbackHost(verifyUri.Host))
        {
            throw new AgentConfigurationException(
                "Verify-endpoint discovery only supports loopback VerifyUrl values. " +
                "For a remote driver, set Driver:BaseUrl and Driver:BearerToken explicitly.");
        }

        _logger.LogInformation("Discovering driver via verify endpoint {VerifyUrl}", verifyUri);

        using var client = _httpClientFactory.CreateClient("driver-unauthenticated");
        client.Timeout = TimeSpan.FromSeconds(_options.RequestTimeoutSeconds);

        HttpResponseMessage response;
        try
        {
            response = await client.GetAsync(verifyUri, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            throw new DriverConnectionException(
                $"Unable to reach verify endpoint '{verifyUri}'. Is the desktop automation driver running?",
                ex);
        }

        using (response)
        {
            if (!response.IsSuccessStatusCode)
            {
                throw new DriverConnectionException(
                    $"Verify endpoint returned HTTP {(int)response.StatusCode}.");
            }

            WebDriverEnvelope<VerifyValueDto>? envelope;
            try
            {
                envelope = await response.Content
                    .ReadFromJsonAsync<WebDriverEnvelope<VerifyValueDto>>(JsonOptions, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (JsonException ex)
            {
                throw new DriverConnectionException("Verify endpoint returned invalid JSON.", ex);
            }

            var value = envelope?.Value
                ?? throw new DriverConnectionException("Verify endpoint response did not contain a value payload.");

            if (string.IsNullOrWhiteSpace(value.Token) && string.IsNullOrWhiteSpace(value.AuthorizationHeader))
                throw new DriverConnectionException("Verify endpoint did not return a token.");

            if (value.Port <= 0)
                throw new DriverConnectionException("Verify endpoint returned an invalid port.");

            var localUser = _currentUserName();
            if (!UsernamesMatch(value.Username, localUser))
            {
                throw new DriverConnectionException(
                    $"Verify endpoint username '{value.Username}' does not match the current Windows user '{localUser}'. " +
                    "Port 9102 may belong to another user on this shared Citrix/RDS host. " +
                    "Configure Driver:BaseUrl and Driver:BearerToken explicitly for your driver instance.");
            }

            var token = !string.IsNullOrWhiteSpace(value.Token)
                ? value.Token
                : ExtractBearerToken(value.AuthorizationHeader);

            if (string.IsNullOrWhiteSpace(token))
                throw new DriverConnectionException("Verify endpoint token was empty after parsing.");

            var baseUri = new Uri($"http://127.0.0.1:{value.Port}/");
            _logger.LogInformation(
                "Discovered driver for user {Username} at {BaseUrl}",
                value.Username,
                $"{baseUri.Scheme}://{baseUri.Authority}");

            return new DriverConnection
            {
                BaseUri = baseUri,
                BearerToken = token,
                DiscoveryMethod = "verify",
                DiscoveredUsername = value.Username
            };
        }
    }

    internal static bool UsernamesMatch(string discovered, string current)
    {
        if (string.IsNullOrWhiteSpace(discovered) || string.IsNullOrWhiteSpace(current))
            return false;

        var left = NormalizeUser(discovered);
        var right = NormalizeUser(current);
        return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeUser(string username)
    {
        var trimmed = username.Trim();
        var slash = trimmed.LastIndexOf('\\');
        if (slash >= 0 && slash < trimmed.Length - 1)
            trimmed = trimmed[(slash + 1)..];

        var at = trimmed.LastIndexOf('@');
        if (at > 0)
            trimmed = trimmed[..at];

        return trimmed;
    }

    private static string ExtractBearerToken(string authorizationHeader)
    {
        const string prefix = "Bearer ";
        if (authorizationHeader.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            return authorizationHeader[prefix.Length..].Trim();

        return authorizationHeader.Trim();
    }

    private static Uri NormalizeBaseUri(Uri uri)
    {
        var builder = new UriBuilder(uri)
        {
            Path = "/",
            Query = string.Empty,
            Fragment = string.Empty
        };
        return builder.Uri;
    }
}
