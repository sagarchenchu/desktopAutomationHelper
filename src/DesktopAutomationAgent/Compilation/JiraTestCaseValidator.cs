using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DesktopAutomationAgent.Configuration;
using Json.Schema;

namespace DesktopAutomationAgent.Compilation;

public sealed class JiraTestCaseValidationResult
{
    public bool IsValid => Errors.Count == 0;

    public string Path { get; init; } = string.Empty;

    public string? Sha256 { get; init; }

    public JiraTestCaseDocument? Document { get; init; }

    public List<string> Errors { get; init; } = [];
}

public sealed class JiraTestCaseValidator
{
    private static readonly Lazy<JsonSchema> Schema = new(LoadSchema);

    public JiraTestCaseValidationResult ValidateFile(string fullPath, string displayPath)
    {
        if (!File.Exists(fullPath))
        {
            return new JiraTestCaseValidationResult
            {
                Path = displayPath,
                Errors = [$"{displayPath}: test-case file not found."]
            };
        }

        byte[] bytes;
        try
        {
            bytes = File.ReadAllBytes(fullPath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new JiraTestCaseValidationResult
            {
                Path = displayPath,
                Errors = [$"{displayPath}: failed to read ({ex.Message})."]
            };
        }

        return ValidateBytes(bytes, displayPath);
    }

    public JiraTestCaseValidationResult ValidateBytes(byte[] bytes, string displayPath)
    {
        var sha = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant();
        JsonNode? node;
        try
        {
            node = JsonNode.Parse(bytes);
        }
        catch (JsonException ex)
        {
            return new JiraTestCaseValidationResult
            {
                Path = displayPath,
                Sha256 = sha,
                Errors = [$"{displayPath}: invalid JSON ({ex.Message})."]
            };
        }

        if (node is null)
        {
            return new JiraTestCaseValidationResult
            {
                Path = displayPath,
                Sha256 = sha,
                Errors = [$"{displayPath}: empty JSON document."]
            };
        }

        var evaluation = Schema.Value.Evaluate(node, new EvaluationOptions
        {
            OutputFormat = OutputFormat.List,
            RequireFormatValidation = true
        });

        var errors = new List<string>();
        if (!evaluation.IsValid)
        {
            foreach (var detail in evaluation.Details.Where(d => d.HasErrors))
            {
                if (detail.Errors is null)
                    continue;
                foreach (var error in detail.Errors)
                {
                    errors.Add($"{displayPath}: schema {detail.InstanceLocation}: {error.Value}");
                }
            }

            if (errors.Count == 0)
                errors.Add($"{displayPath}: failed JSON Schema validation.");
        }

        JiraTestCaseDocument? document = null;
        try
        {
            document = JsonSerializer.Deserialize<JiraTestCaseDocument>(bytes, CompilationJson.DocumentRead);
        }
        catch (Exception ex) when (ex is JsonException or NotSupportedException)
        {
            errors.Add($"{displayPath}: failed to deserialize ({ex.Message}).");
        }

        if (document is not null)
        {
            if (document.ExtensionData is { Count: > 0 })
            {
                foreach (var key in document.ExtensionData.Keys.OrderBy(k => k, StringComparer.Ordinal))
                    errors.Add($"{displayPath}: unknown property '{key}'.");
            }

            if (!JiraKeyContract.IsCanonical(document.JiraKey))
                errors.Add($"{displayPath}: jiraKey '{document.JiraKey}' is not canonical.");

            if (string.IsNullOrWhiteSpace(document.Feature) && (document.Scenarios is null || document.Scenarios.Count == 0))
                errors.Add($"{displayPath}: either feature or scenarios is required.");

            if (string.IsNullOrWhiteSpace(document.ResolveRevisionToken()))
                errors.Add($"{displayPath}: revision or updatedAt is required.");

            // Reject credential-like keys in tags/testDataRefs by convention.
            if (document.TestDataRefs is not null)
            {
                foreach (var refer in document.TestDataRefs)
                {
                    if (refer.Contains("password", StringComparison.OrdinalIgnoreCase)
                        || refer.Contains("secret", StringComparison.OrdinalIgnoreCase)
                        || refer.Contains("token", StringComparison.OrdinalIgnoreCase)
                        || refer.Contains("credential", StringComparison.OrdinalIgnoreCase))
                    {
                        errors.Add($"{displayPath}: testDataRefs must not reference secrets ('{refer}').");
                    }
                }
            }
        }

        return new JiraTestCaseValidationResult
        {
            Path = displayPath,
            Sha256 = sha,
            Document = errors.Count == 0 ? document : null,
            Errors = errors
        };
    }

    private static JsonSchema LoadSchema()
    {
        var resource = "DesktopAutomationAgent.Compilation.jira-test-case.schema.json";
        using var stream = typeof(JiraTestCaseValidator).Assembly.GetManifestResourceStream(resource)
            ?? throw new InvalidOperationException($"Missing embedded schema '{resource}'.");
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return JsonSchema.FromText(reader.ReadToEnd());
    }
}
