using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Driver.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Driver;

public sealed class DriverCatalogClient : IDriverCatalogClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly DriverOptions _options;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<DriverCatalogClient> _logger;

    public DriverCatalogClient(
        IOptions<AgentOptions> options,
        IHttpClientFactory httpClientFactory,
        ILogger<DriverCatalogClient> logger)
    {
        _options = options.Value.Driver;
        _httpClientFactory = httpClientFactory;
        _logger = logger;
    }

    public async Task<StatusValueDto> GetStatusAsync(
        DriverConnection connection,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(connection);

        using var response = await SendAuthenticatedAsync(
            connection,
            "status",
            cancellationToken).ConfigureAwait(false);

        if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
        {
            throw new DriverCatalogException(
                "Driver authentication failed for GET /status (HTTP 401). Check the bearer token.");
        }

        if (!response.IsSuccessStatusCode)
        {
            throw new DriverCatalogException(
                $"GET /status returned HTTP {(int)response.StatusCode}.");
        }

        WebDriverEnvelope<StatusValueDto>? envelope;
        try
        {
            envelope = await response.Content
                .ReadFromJsonAsync<WebDriverEnvelope<StatusValueDto>>(JsonOptions, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (JsonException ex)
        {
            throw new DriverCatalogException("GET /status returned invalid JSON.", ex);
        }

        var value = envelope?.Value
            ?? throw new DriverCatalogException("GET /status response did not contain a value payload.");

        _logger.LogInformation(
            "Driver status ready={Ready} version={Version}",
            value.Ready,
            value.Build?.Version ?? "(unknown)");

        return value;
    }

    public async Task<OperationsCatalogDto> GetOperationsAsync(
        DriverConnection connection,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(connection);

        using var response = await SendAuthenticatedAsync(
            connection,
            "ui/operations",
            cancellationToken).ConfigureAwait(false);

        if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
        {
            throw new DriverCatalogException(
                "Driver authentication failed for GET /ui/operations (HTTP 401). Check the bearer token.");
        }

        if (!response.IsSuccessStatusCode)
        {
            throw new DriverCatalogException(
                $"GET /ui/operations returned HTTP {(int)response.StatusCode}.");
        }

        UiEnvelope<OperationsCatalogDto>? envelope;
        try
        {
            envelope = await response.Content
                .ReadFromJsonAsync<UiEnvelope<OperationsCatalogDto>>(JsonOptions, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (JsonException ex)
        {
            throw new DriverCatalogException("GET /ui/operations returned invalid JSON.", ex);
        }

        if (envelope is null)
            throw new DriverCatalogException("GET /ui/operations response was empty.");

        if (!envelope.Success)
        {
            throw new DriverCatalogException(
                $"GET /ui/operations failed: {SecretRedactor.Redact(envelope.Error) ?? "unknown error"}");
        }

        var catalog = envelope.Value
            ?? throw new DriverCatalogException("GET /ui/operations response did not contain a catalog value.");

        _logger.LogInformation(
            "Loaded operation catalog schemaVersion={SchemaVersion} driverVersion={DriverVersion} count={Count}",
            catalog.SchemaVersion,
            catalog.DriverVersion,
            catalog.Operations.Count);

        return catalog;
    }

    private async Task<HttpResponseMessage> SendAuthenticatedAsync(
        DriverConnection connection,
        string relativePath,
        CancellationToken cancellationToken)
    {
        using var client = _httpClientFactory.CreateClient("driver-authenticated");
        client.BaseAddress = connection.BaseUri;
        client.Timeout = TimeSpan.FromSeconds(_options.RequestTimeoutSeconds);
        client.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", connection.BearerToken);

        try
        {
            return await client.GetAsync(relativePath, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            throw new DriverConnectionException(
                $"Unable to reach driver at {connection.SafeBaseUrl} for '{relativePath}'.",
                ex);
        }
    }
}
