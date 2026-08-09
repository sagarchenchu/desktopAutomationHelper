using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Configuration;
using DesktopAutomationAgent.Plans;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace DesktopAutomationAgent.Driver;

public sealed class DriverUiClient : IDriverUiClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly RunnerOptions _runnerOptions;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<DriverUiClient> _logger;

    public DriverUiClient(
        IOptions<AgentOptions> options,
        IHttpClientFactory httpClientFactory,
        ILogger<DriverUiClient> logger)
    {
        _runnerOptions = options.Value.Runner;
        _httpClientFactory = httpClientFactory;
        _logger = logger;
    }

    public async Task<UiExecutionResponse> ExecuteStepAsync(
        DriverConnection connection,
        PlanStep step,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(connection);
        ArgumentNullException.ThrowIfNull(step);

        using var client = _httpClientFactory.CreateClient("driver-authenticated");
        client.BaseAddress = connection.BaseUri;
        client.Timeout = TimeSpan.FromSeconds(_runnerOptions.StepTransportTimeoutSeconds);
        client.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", connection.BearerToken);

        var payload = BuildFlattenedPayload(step);
        var json = JsonSerializer.Serialize(payload);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");

        HttpResponseMessage response;
        try
        {
            response = await client.PostAsync("ui", content, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException ex) when (cancellationToken.IsCancellationRequested)
        {
            throw new UiExecutionException(
                UiFailureClassification.Cancelled,
                $"Step '{step.Id}' was cancelled.",
                innerException: ex);
        }
        catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
        {
            throw new UiExecutionException(
                UiFailureClassification.ExecutionTimeout,
                $"Step '{step.Id}' timed out after {_runnerOptions.StepTransportTimeoutSeconds} seconds.",
                innerException: ex);
        }
        catch (Exception ex) when (ex is HttpRequestException)
        {
            throw new UiExecutionException(
                UiFailureClassification.DriverUnavailable,
                $"Unable to reach driver at {connection.SafeBaseUrl} for step '{step.Id}'.",
                innerException: ex);
        }

        using (response)
        {
            if (response.StatusCode == HttpStatusCode.Unauthorized)
            {
                throw new UiExecutionException(
                    UiFailureClassification.Authentication,
                    $"Driver authentication failed for step '{step.Id}' (HTTP 401).");
            }

            var bodyBytes = await ReadLimitedBodyAsync(response, cancellationToken).ConfigureAwait(false);
            UiExecutionResponse? envelope;
            try
            {
                envelope = JsonSerializer.Deserialize<UiExecutionResponse>(bodyBytes, JsonOptions);
            }
            catch (JsonException ex)
            {
                throw new UiExecutionException(
                    UiFailureClassification.OperationFailure,
                    $"Step '{step.Id}' returned invalid JSON.",
                    innerException: ex);
            }

            if (envelope is null)
            {
                throw new UiExecutionException(
                    UiFailureClassification.OperationFailure,
                    $"Step '{step.Id}' returned an empty response.");
            }

            if (!response.IsSuccessStatusCode && envelope.Success)
            {
                throw new UiExecutionException(
                    UiFailureClassification.OperationFailure,
                    $"Step '{step.Id}' returned HTTP {(int)response.StatusCode}.",
                    envelope);
            }

            if (!envelope.Success)
            {
                var detail = SecretRedactor.Redact(envelope.Error ?? envelope.Reason ?? "unknown error");
                throw new UiExecutionException(
                    UiFailureClassification.OperationFailure,
                    $"Step '{step.Id}' failed: {detail}",
                    envelope);
            }

            _logger.LogInformation(
                "Step {StepId} operation {Operation} succeeded.",
                step.Id,
                step.Operation);

            return envelope;
        }
    }

    private async Task<byte[]> ReadLimitedBodyAsync(
        HttpResponseMessage response,
        CancellationToken cancellationToken)
    {
        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        using var memory = new MemoryStream();
        var buffer = new byte[8192];
        var total = 0;
        int read;

        while ((read = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(false)) > 0)
        {
            total += read;
            if (total > _runnerOptions.MaxResponseBytes)
            {
                throw new UiExecutionException(
                    UiFailureClassification.OperationFailure,
                    $"Driver response exceeded maximum size of {_runnerOptions.MaxResponseBytes} bytes.");
            }

            memory.Write(buffer, 0, read);
        }

        return memory.ToArray();
    }

    internal static Dictionary<string, JsonElement> BuildFlattenedPayload(PlanStep step)
    {
        var payload = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
        {
            ["operation"] = JsonSerializer.SerializeToElement(step.Operation)
        };

        if (step.Arguments is not null)
        {
            foreach (var pair in step.Arguments)
            {
                payload[pair.Key] = pair.Value;
            }
        }

        return payload;
    }
}
