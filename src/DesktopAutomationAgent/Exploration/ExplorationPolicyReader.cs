using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Driver;
using Json.Schema;

namespace DesktopAutomationAgent.Exploration;

public sealed class ExplorationPolicyValidationResult
{
    public bool IsValid => Errors.Count == 0;

    public List<string> Errors { get; } = [];

    public List<string> Warnings { get; } = [];

    public ExplorationPolicyDocument? Policy { get; set; }

    public string? FilePath { get; set; }

    public string? Sha256 { get; set; }
}

public sealed class ExplorationPolicyReader
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private static readonly Lazy<JsonSchema> Schema = new(LoadEmbeddedSchema);

    private static readonly HashSet<string> AllowedModes = new(StringComparer.OrdinalIgnoreCase)
    {
        "observe",
        "navigate"
    };

    private static readonly HashSet<string> AllowedCategories = new(StringComparer.Ordinal)
    {
        "expandCollapse",
        "tabSelect",
        "menuOpen",
        "scroll",
        "focus",
        "hover",
        "selectItem",
        "invoke"
    };

    public ExplorationPolicyValidationResult Read(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return new ExplorationPolicyValidationResult
            {
                Errors = { "Policy path is required." }
            };
        }

        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            return new ExplorationPolicyValidationResult
            {
                FilePath = fullPath,
                Errors = { $"Policy file not found: {fullPath}" }
            };
        }

        string json;
        try
        {
            json = File.ReadAllText(fullPath);
        }
        catch (Exception ex)
        {
            return new ExplorationPolicyValidationResult
            {
                FilePath = fullPath,
                Errors = { $"Unable to read policy file: {SecretRedactor.Redact(ex.Message)}" }
            };
        }

        var schemaResult = ValidateJsonAgainstSchema(json);
        schemaResult.FilePath = fullPath;
        schemaResult.Sha256 = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(json))).ToLowerInvariant();
        if (!schemaResult.IsValid)
            return schemaResult;

        ExplorationPolicyDocument? policy;
        try
        {
            policy = JsonSerializer.Deserialize<ExplorationPolicyDocument>(json, JsonOptions);
        }
        catch (JsonException ex)
        {
            return new ExplorationPolicyValidationResult
            {
                FilePath = fullPath,
                Sha256 = schemaResult.Sha256,
                Errors = { $"Policy JSON is invalid: {SecretRedactor.Redact(ex.Message)}" }
            };
        }

        if (policy is null)
        {
            return new ExplorationPolicyValidationResult
            {
                FilePath = fullPath,
                Sha256 = schemaResult.Sha256,
                Errors = { "Policy JSON deserialized to null." }
            };
        }

        var result = Validate(policy);
        result.FilePath = fullPath;
        result.Sha256 = schemaResult.Sha256;
        return result;
    }

    /// <summary>
    /// Validates raw policy JSON against exploration-policy.schema.json so unknown
    /// properties (e.g. misspelled maxActions) cannot silently default.
    /// </summary>
    public ExplorationPolicyValidationResult ValidateJsonAgainstSchema(string json)
    {
        var result = new ExplorationPolicyValidationResult();
        JsonNode? node;
        try
        {
            node = JsonNode.Parse(json);
        }
        catch (JsonException ex)
        {
            result.Errors.Add($"Policy JSON is invalid: {SecretRedactor.Redact(ex.Message)}");
            return result;
        }

        if (node is null)
        {
            result.Errors.Add("Policy JSON parsed to null.");
            return result;
        }

        EvaluationResults evaluation;
        try
        {
            evaluation = Schema.Value.Evaluate(node, new EvaluationOptions
            {
                OutputFormat = OutputFormat.List,
                RequireFormatValidation = true
            });
        }
        catch (Exception ex)
        {
            result.Errors.Add($"Policy schema evaluation failed: {SecretRedactor.Redact(ex.Message)}");
            return result;
        }

        if (evaluation.IsValid)
            return result;

        foreach (var detail in evaluation.Details)
        {
            if (detail.Errors is null || detail.Errors.Count == 0)
                continue;

            var location = detail.InstanceLocation.ToString();
            foreach (var error in detail.Errors)
            {
                result.Errors.Add($"schema {location}: {error.Key} — {error.Value}");
            }
        }

        if (result.Errors.Count == 0)
            result.Errors.Add("Policy JSON failed schema validation.");

        return result;
    }

    public ExplorationPolicyValidationResult Validate(ExplorationPolicyDocument policy)
    {
        ArgumentNullException.ThrowIfNull(policy);
        var result = new ExplorationPolicyValidationResult { Policy = policy };

        if (policy.PolicySchemaVersion != ExplorationPolicyDocument.SchemaVersion)
            result.Errors.Add($"Unsupported policySchemaVersion '{policy.PolicySchemaVersion}'.");

        if (!AllowedModes.Contains(policy.Mode))
            result.Errors.Add("mode must be 'observe' or 'navigate'.");

        if (string.IsNullOrWhiteSpace(policy.Application.ProcessName))
            result.Errors.Add("application.processName is required.");

        if (!string.IsNullOrWhiteSpace(policy.Application.ExecutableSha256)
            && policy.Application.ExecutableSha256.Length != 64)
        {
            result.Errors.Add("application.executableSha256 must be a 64-character hex digest when set.");
        }

        if (policy.AllowedActionCategories.Count > 16)
            result.Errors.Add("allowedActionCategories exceeds maxItems 16.");

        if (policy.AllowRules.Count > 500)
            result.Errors.Add("allowRules exceeds maxItems 500.");

        if (policy.DenyRules.Count > 500)
            result.Errors.Add("denyRules exceeds maxItems 500.");

        foreach (var category in policy.AllowedActionCategories)
        {
            if (!AllowedCategories.Contains(category))
                result.Errors.Add($"Unknown allowedActionCategory '{category}'.");
        }

        if (string.Equals(policy.Mode, "navigate", StringComparison.OrdinalIgnoreCase)
            && policy.AllowedActionCategories.Count == 0
            && policy.AllowRules.Count == 0)
        {
            result.Warnings.Add("navigate mode has no allowedActionCategories or allowRules; deny-by-default will permit nothing.");
        }

        if (policy.AllowHover && !policy.AllowedActionCategories.Contains("hover", StringComparer.Ordinal))
            result.Warnings.Add("allowHover is true but 'hover' is not in allowedActionCategories.");

        if (policy.AllowCoordinateFallback)
            result.Errors.Add("allowCoordinateFallback must be false in Phase 4B.");

        if (policy.AllowDoubleClick)
            result.Errors.Add("allowDoubleClick is unsupported in Phase 4B.");

        ValidateBudgets(policy.Budgets, result);
        ValidateRules(policy.AllowRules, "allowRules", result);
        ValidateRules(policy.DenyRules, "denyRules", result);

        return result;
    }

    private static void ValidateBudgets(ExplorationBudgets budgets, ExplorationPolicyValidationResult result)
    {
        if (budgets.MaxStates is < 1 or > 10000)
            result.Errors.Add("budgets.maxStates must be between 1 and 10000.");
        if (budgets.MaxActions is < 0 or > 100000)
            result.Errors.Add("budgets.maxActions must be between 0 and 100000.");
        if (budgets.MaxDepth is < 0 or > 1000)
            result.Errors.Add("budgets.maxDepth must be between 0 and 1000.");
        if (budgets.TotalDurationSeconds is < 1 or > 86400)
            result.Errors.Add("budgets.totalDurationSeconds must be between 1 and 86400.");
        if (budgets.PerActionTimeoutSeconds is < 1 or > 600)
            result.Errors.Add("budgets.perActionTimeoutSeconds must be between 1 and 600.");
        if (budgets.StabilizationTimeoutSeconds is < 1 or > 600)
            result.Errors.Add("budgets.stabilizationTimeoutSeconds must be between 1 and 600.");
        if (budgets.StabilizationQuietPeriodMilliseconds is < 0 or > 60000)
            result.Errors.Add("budgets.stabilizationQuietPeriodMilliseconds must be between 0 and 60000.");
    }

    private static void ValidateRules(
        IEnumerable<ExplorationElementRule> rules,
        string location,
        ExplorationPolicyValidationResult result)
    {
        var index = 0;
        foreach (var rule in rules)
        {
            index++;
            var prefix = $"{location}[{index}]";
            if (string.IsNullOrWhiteSpace(rule.AutomationId)
                && string.IsNullOrWhiteSpace(rule.Name)
                && string.IsNullOrWhiteSpace(rule.NameContains)
                && string.IsNullOrWhiteSpace(rule.ControlType)
                && string.IsNullOrWhiteSpace(rule.ClassName))
            {
                result.Errors.Add($"{prefix}: at least one match field is required.");
            }

            if (!string.IsNullOrWhiteSpace(rule.ActionCategory)
                && rule.ActionCategory != "*"
                && !AllowedCategories.Contains(rule.ActionCategory))
            {
                result.Errors.Add($"{prefix}: unknown actionCategory '{rule.ActionCategory}'.");
            }
        }
    }

    private static JsonSchema LoadEmbeddedSchema()
    {
        var assembly = typeof(ExplorationPolicyReader).Assembly;
        const string resourceName = "DesktopAutomationAgent.Exploration.exploration-policy.schema.json";
        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream is null)
        {
            // Fallback for local/dev when embedding path differs.
            var candidates = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "exploration-policy.schema.json"),
                Path.Combine(AppContext.BaseDirectory, "schemas", "exploration-policy.schema.json"),
                Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "automation", "schemas", "exploration-policy.schema.json"))
            };
            foreach (var candidate in candidates)
            {
                if (File.Exists(candidate))
                    return JsonSchema.FromText(File.ReadAllText(candidate));
            }

            throw new InvalidOperationException(
                $"Embedded exploration policy schema '{resourceName}' was not found.");
        }

        using var reader = new StreamReader(stream, Encoding.UTF8);
        return JsonSchema.FromText(reader.ReadToEnd());
    }
}
