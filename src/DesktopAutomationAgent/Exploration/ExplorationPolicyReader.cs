using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Driver;

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
                Errors = { $"Policy JSON is invalid: {SecretRedactor.Redact(ex.Message)}" }
            };
        }

        if (policy is null)
        {
            return new ExplorationPolicyValidationResult
            {
                FilePath = fullPath,
                Errors = { "Policy JSON deserialized to null." }
            };
        }

        var result = Validate(policy);
        result.FilePath = fullPath;
        result.Sha256 = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(json))).ToLowerInvariant();
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
        if (budgets.MaxStates < 1)
            result.Errors.Add("budgets.maxStates must be >= 1.");
        if (budgets.MaxActions < 0)
            result.Errors.Add("budgets.maxActions must be >= 0.");
        if (budgets.MaxDepth < 0)
            result.Errors.Add("budgets.maxDepth must be >= 0.");
        if (budgets.TotalDurationSeconds < 1)
            result.Errors.Add("budgets.totalDurationSeconds must be >= 1.");
        if (budgets.PerActionTimeoutSeconds < 1)
            result.Errors.Add("budgets.perActionTimeoutSeconds must be >= 1.");
        if (budgets.StabilizationTimeoutSeconds < 1)
            result.Errors.Add("budgets.stabilizationTimeoutSeconds must be >= 1.");
        if (budgets.StabilizationQuietPeriodMilliseconds < 0)
            result.Errors.Add("budgets.stabilizationQuietPeriodMilliseconds must be >= 0.");
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
}
