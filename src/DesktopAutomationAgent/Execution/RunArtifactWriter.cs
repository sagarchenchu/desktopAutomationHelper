using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Driver;

namespace DesktopAutomationAgent.Execution;

public sealed class RunArtifactWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    private static readonly HashSet<string> SensitivePropertyNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "password",
        "token",
        "bearerToken",
        "authorization",
        "authorizationHeader",
        "secret",
        "apiKey",
        "accessToken",
        "refreshToken",
        "credential",
        "credentials"
    };

    public void WriteRunReport(string runDirectory, RunReport report)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(runDirectory);
        ArgumentNullException.ThrowIfNull(report);

        Directory.CreateDirectory(runDirectory);
        var targetPath = Path.Combine(runDirectory, "run.json");
        var tempPath = targetPath + ".tmp";

        var redacted = RedactReport(report);
        var json = JsonSerializer.Serialize(redacted, JsonOptions);
        File.WriteAllText(tempPath, json);
        File.Move(tempPath, targetPath, overwrite: true);
    }

    internal static RunReport RedactReport(RunReport report) =>
        new()
        {
            RunId = report.RunId,
            Status = report.Status,
            ExitCode = report.ExitCode,
            PlanPath = report.PlanPath,
            PlanId = report.PlanId,
            PlanName = report.PlanName,
            PlanSha256 = report.PlanSha256,
            DryRun = report.DryRun,
            DriverBaseUrl = report.DriverBaseUrl,
            CatalogSchemaVersion = report.CatalogSchemaVersion,
            StartedAtUtc = report.StartedAtUtc,
            FinishedAtUtc = report.FinishedAtUtc,
            Steps = report.Steps.Select(RedactStep).ToArray(),
            OnFailureSteps = report.OnFailureSteps.Select(RedactStep).ToArray(),
            Failure = report.Failure is null
                ? null
                : new RunFailure
                {
                    Classification = report.Failure.Classification,
                    Message = SecretRedactor.Redact(report.Failure.Message),
                    StepId = report.Failure.StepId,
                    ScreenshotPath = report.Failure.ScreenshotPath
                }
        };

    private static StepRunResult RedactStep(StepRunResult step) =>
        new()
        {
            Id = step.Id,
            Operation = step.Operation,
            Phase = step.Phase,
            Success = step.Success,
            Sensitive = step.Sensitive,
            CaptureResponse = step.CaptureResponse,
            Skipped = step.Skipped,
            SkipReason = step.SkipReason,
            Arguments = step.Sensitive ? null : RedactJsonMap(step.Arguments),
            ResponseValue = step.CaptureResponse ? RedactJsonElement(step.ResponseValue) : null,
            Error = SecretRedactor.Redact(step.Error),
            ScreenshotPath = step.ScreenshotPath,
            Assertions = step.Assertions.Select(assertion => RedactAssertion(assertion, step.Sensitive)).ToArray(),
            Duration = step.Duration
        };

    private static AssertionRunResult RedactAssertion(AssertionRunResult assertion, bool stepSensitive) =>
        new()
        {
            Path = assertion.Path,
            Operator = assertion.Operator,
            Passed = assertion.Passed,
            Message = assertion.Message,
            Expected = stepSensitive ? null : RedactJsonElement(assertion.Expected),
            Actual = stepSensitive ? null : RedactJsonElement(assertion.Actual)
        };

    private static Dictionary<string, JsonElement>? RedactJsonMap(Dictionary<string, JsonElement>? map)
    {
        if (map is null)
            return null;

        var result = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (var pair in map)
        {
            result[pair.Key] = IsSensitivePropertyName(pair.Key)
                ? JsonSerializer.SerializeToElement("[REDACTED]")
                : RedactJsonElement(pair.Value) ?? pair.Value;
        }

        return result;
    }

    private static JsonElement? RedactJsonElement(JsonElement? element)
    {
        if (element is null)
            return null;

        return RedactJsonElement(element.Value);
    }

    private static JsonElement? RedactJsonElement(JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
            {
                var obj = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
                foreach (var property in element.EnumerateObject())
                {
                    obj[property.Name] = IsSensitivePropertyName(property.Name)
                        ? JsonSerializer.SerializeToElement("[REDACTED]")
                        : RedactJsonElement(property.Value) ?? property.Value;
                }

                return JsonSerializer.SerializeToElement(obj);
            }
            case JsonValueKind.Array:
            {
                var items = element.EnumerateArray()
                    .Select(item => RedactJsonElement(item) ?? item)
                    .ToArray();
                return JsonSerializer.SerializeToElement(items);
            }
            default:
                return element;
        }
    }

    private static bool IsSensitivePropertyName(string name) =>
        SensitivePropertyNames.Contains(name);
}
