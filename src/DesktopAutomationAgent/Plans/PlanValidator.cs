using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Plans;

public sealed class PlanValidator
{
    private static readonly Regex PlanIdPattern = new(
        @"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly HashSet<string> AllowedAssertionOperators = new(StringComparer.Ordinal)
    {
        "equals",
        "notEquals",
        "contains",
        "matchesRegex",
        "isTrue",
        "isFalse",
        "isNull",
        "isNotNull"
    };

    private static readonly HashSet<string> ReservedArgumentNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "operation",
        "authorization",
        "bearerToken"
    };

    private const int MaxCombinedSteps = 1000;

    public PlanValidationResult Validate(PlanManifest manifest, string relativePath)
    {
        ArgumentNullException.ThrowIfNull(manifest);

        var errors = new List<string>();
        var stepCount = manifest.Steps?.Count ?? 0;
        var onFailureCount = manifest.OnFailureSteps?.Count ?? 0;

        if (manifest.SchemaVersion != 1)
        {
            errors.Add($"{relativePath}: unsupported schemaVersion {manifest.SchemaVersion}; expected 1.");
        }

        if (manifest.CatalogSchemaVersion != 2)
        {
            errors.Add(
                $"{relativePath}: unsupported catalogSchemaVersion {manifest.CatalogSchemaVersion}; expected 2.");
        }

        if (string.IsNullOrWhiteSpace(manifest.PlanId) || !PlanIdPattern.IsMatch(manifest.PlanId))
        {
            errors.Add($"{relativePath}: planId must match ^[A-Za-z0-9][A-Za-z0-9._-]{{0,127}}$.");
        }

        if (string.IsNullOrWhiteSpace(manifest.Name))
        {
            errors.Add($"{relativePath}: name is required.");
        }

        if (manifest.ExtensionData is { Count: > 0 })
        {
            foreach (var key in manifest.ExtensionData.Keys)
            {
                errors.Add($"{relativePath}: unknown top-level property '{key}'.");
            }
        }

        if (manifest.Steps is null)
        {
            errors.Add($"{relativePath}: steps is required.");
        }
        else if (manifest.Steps.Count == 0)
        {
            errors.Add($"{relativePath}: steps must contain at least one step.");
        }

        if (stepCount + onFailureCount > MaxCombinedSteps)
        {
            errors.Add(
                $"{relativePath}: combined steps and onFailureSteps count {stepCount + onFailureCount} exceeds maximum {MaxCombinedSteps}.");
        }

        var seenIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        ValidateStepList(manifest.Steps, "steps", relativePath, errors, seenIds, allowAssertions: true);
        ValidateStepList(manifest.OnFailureSteps, "onFailureSteps", relativePath, errors, seenIds, allowAssertions: false);

        return new PlanValidationResult
        {
            PlanPath = relativePath,
            PlanId = manifest.PlanId ?? string.Empty,
            Name = manifest.Name ?? string.Empty,
            StepCount = stepCount,
            OnFailureStepCount = onFailureCount,
            CleanupStepCount = 0,
            TotalStepCount = stepCount + onFailureCount,
            Errors = errors
        };
    }

    private static void ValidateStepList(
        List<PlanStep>? steps,
        string listName,
        string relativePath,
        List<string> errors,
        Dictionary<string, string> seenIds,
        bool allowAssertions)
    {
        if (steps is null)
            return;

        for (var i = 0; i < steps.Count; i++)
        {
            var step = steps[i];
            var location = $"{relativePath}: {listName}[{i}]";

            if (step.ExtensionData is { Count: > 0 })
            {
                foreach (var key in step.ExtensionData.Keys)
                {
                    errors.Add($"{location}: unknown property '{key}'.");
                }
            }

            if (string.IsNullOrWhiteSpace(step.Id))
            {
                errors.Add($"{location}: id is required.");
            }
            else if (seenIds.TryGetValue(step.Id, out var firstLocation))
            {
                errors.Add($"{location}: id duplicates {firstLocation}.id.");
            }
            else
            {
                seenIds[step.Id] = $"{listName}[{i}]";
            }

            if (string.IsNullOrWhiteSpace(step.Operation))
            {
                errors.Add($"{location}: operation is required.");
            }
            else if (!string.Equals(step.Operation, step.Operation.Trim(), StringComparison.Ordinal))
            {
                errors.Add($"{location}: operation must not contain leading or trailing whitespace.");
            }

            if (step.Arguments is null)
            {
                errors.Add($"{location}: arguments is required.");
            }
            else
            {
                foreach (var argumentName in step.Arguments.Keys)
                {
                    if (ReservedArgumentNames.Contains(argumentName))
                    {
                        errors.Add($"{location}: arguments must not use reserved name '{argumentName}'.");
                    }
                }
            }

            if (!allowAssertions)
            {
                if (step.Assertions is not null)
                {
                    errors.Add($"{location}: assertions are not allowed on cleanup steps.");
                }

                if (step.CaptureResponse is not null)
                {
                    errors.Add($"{location}: captureResponse is not allowed on cleanup steps.");
                }
            }
            else
            {
                ValidateAssertions(step.Assertions, location, errors);
            }
        }
    }

    private static void ValidateAssertions(
        List<PlanAssertion>? assertions,
        string stepLocation,
        List<string> errors)
    {
        if (assertions is null)
            return;

        for (var i = 0; i < assertions.Count; i++)
        {
            var assertion = assertions[i];
            var location = $"{stepLocation}.assertions[{i}]";

            if (assertion.ExtensionData is { Count: > 0 })
            {
                foreach (var key in assertion.ExtensionData.Keys)
                {
                    errors.Add($"{location}: unknown property '{key}'.");
                }
            }

            var pointer = assertion.Path ?? string.Empty;
            if (!IsValidJsonPointer(pointer))
            {
                errors.Add($"{location}: path '{pointer}' is not a valid JSON pointer.");
            }

            if (string.IsNullOrWhiteSpace(assertion.Operator))
            {
                errors.Add($"{location}: operator is required.");
                continue;
            }

            if (!AllowedAssertionOperators.Contains(assertion.Operator))
            {
                errors.Add($"{location}: operator '{assertion.Operator}' is not supported.");
                continue;
            }

            if (RequiresExpected(assertion.Operator) && assertion.Expected is null)
            {
                errors.Add($"{location}: expected is required for operator '{assertion.Operator}'.");
            }
        }
    }

    private static bool RequiresExpected(string op) =>
        op is "equals" or "notEquals" or "contains" or "matchesRegex";

    internal static bool IsValidJsonPointer(string path)
    {
        if (path.Length == 0)
            return true;

        if (path[0] != '/')
            return false;

        for (var i = 1; i < path.Length; i++)
        {
            if (path[i] != '~')
                continue;

            if (i + 1 >= path.Length)
                return false;

            if (path[i + 1] is not ('0' or '1'))
                return false;

            i++;
        }

        return true;
    }
}
