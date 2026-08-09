using System.Text.Json;
using DesktopAutomationAgent.Driver.Models;

namespace DesktopAutomationAgent.Plans;

public sealed class PlanCatalogPreflight
{
    private readonly Dictionary<string, OperationDescriptorDto> _canonicalByName;
    private readonly Dictionary<string, string> _aliasOwners;

    public PlanCatalogPreflight(OperationsCatalogDto catalog)
    {
        ArgumentNullException.ThrowIfNull(catalog);

        _canonicalByName = new Dictionary<string, OperationDescriptorDto>(StringComparer.OrdinalIgnoreCase);
        _aliasOwners = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var operation in catalog.Operations)
        {
            if (string.IsNullOrWhiteSpace(operation.Name))
                continue;

            _canonicalByName[operation.Name] = operation;

            foreach (var alias in operation.Aliases)
            {
                if (!string.IsNullOrWhiteSpace(alias))
                    _aliasOwners[alias] = operation.Name;
            }

            foreach (var alias in operation.DeprecatedAliases)
            {
                if (!string.IsNullOrWhiteSpace(alias))
                    _aliasOwners[alias] = operation.Name;
            }
        }
    }

    public IReadOnlyList<string> Validate(PlanManifest manifest, string relativePath)
    {
        ArgumentNullException.ThrowIfNull(manifest);

        var errors = new List<string>();
        ValidateStepList(manifest.Steps, "steps", relativePath, errors);
        ValidateStepList(manifest.OnFailureSteps, "onFailureSteps", relativePath, errors);
        ValidateSessionLifecycle(manifest, relativePath, errors);
        return errors;
    }

    private void ValidateStepList(
        List<PlanStep>? steps,
        string listName,
        string relativePath,
        List<string> errors)
    {
        if (steps is null)
            return;

        for (var i = 0; i < steps.Count; i++)
        {
            var step = steps[i];
            var location = $"{relativePath}: {listName}[{i}]";
            ValidateStep(step, location, errors);
        }
    }

    private void ValidateStep(PlanStep step, string location, List<string> errors)
    {
        var operation = step.Operation;
        if (string.IsNullOrWhiteSpace(operation))
            return;

        if (_aliasOwners.TryGetValue(operation, out var canonicalFromAlias))
        {
            errors.Add(
                $"{location}: operation '{operation}' is an alias; use canonical operation '{canonicalFromAlias}'.");
            return;
        }

        if (!_canonicalByName.TryGetValue(operation, out var descriptor))
        {
            errors.Add($"{location}: unknown operation '{operation}'.");
            return;
        }

        // Normalize in memory only; never rewrite the plan file.
        step.Operation = descriptor.Name;

        if (descriptor.Deprecated)
        {
            errors.Add($"{location}: operation '{descriptor.Name}' is deprecated.");
        }

        ValidateRequiredInputs(step, descriptor, location, errors);
    }

    private static void ValidateRequiredInputs(
        PlanStep step,
        OperationDescriptorDto descriptor,
        string location,
        List<string> errors)
    {
        var arguments = step.Arguments ?? new Dictionary<string, JsonElement>();

        foreach (var required in descriptor.RequiredInputs)
        {
            if (!TryGetArgument(arguments, required, out var value, out _))
            {
                errors.Add($"{location}: missing required argument '{required}'.");
                continue;
            }

            if (!IsPresentArgumentValue(required, value))
            {
                errors.Add($"{location}: required argument '{required}' must not be null, blank, or an empty locator object.");
            }
        }

        if (descriptor.RequiredInputAlternatives.Count == 0)
            return;

        if (!descriptor.RequiredInputAlternatives.Any(alternative => IsCompleteAlternative(arguments, alternative)))
        {
            var alternatives = string.Join(
                " | ",
                descriptor.RequiredInputAlternatives.Select(group => $"[{string.Join(", ", group)}]"));
            errors.Add(
                $"{location}: must provide one complete argument alternative: {alternatives}.");
        }
    }

    private static bool IsCompleteAlternative(
        IReadOnlyDictionary<string, JsonElement> arguments,
        IReadOnlyList<string> alternative)
    {
        foreach (var required in alternative)
        {
            if (!TryGetArgument(arguments, required, out var value, out _))
                return false;

            if (!IsPresentArgumentValue(required, value))
                return false;
        }

        return true;
    }

    private static bool TryGetArgument(
        IReadOnlyDictionary<string, JsonElement> arguments,
        string name,
        out JsonElement value,
        out string matchedKey)
    {
        if (arguments.TryGetValue(name, out value))
        {
            matchedKey = name;
            return true;
        }

        foreach (var pair in arguments)
        {
            if (!string.Equals(pair.Key, name, StringComparison.OrdinalIgnoreCase))
                continue;

            value = pair.Value;
            matchedKey = pair.Key;
            return true;
        }

        value = default;
        matchedKey = string.Empty;
        return false;
    }

    private static bool IsPresentArgumentValue(string argumentName, JsonElement value) =>
        value.ValueKind switch
        {
            JsonValueKind.Null => false,
            JsonValueKind.Undefined => false,
            JsonValueKind.String => !string.IsNullOrWhiteSpace(value.GetString()),
            JsonValueKind.Object when IsLocatorArgumentName(argumentName) => value.EnumerateObject().Any(),
            _ => true
        };

    private static bool IsLocatorArgumentName(string argumentName) =>
        string.Equals(argumentName, "locator", StringComparison.OrdinalIgnoreCase)
        || argumentName.EndsWith("Locator", StringComparison.OrdinalIgnoreCase);

    private void ValidateSessionLifecycle(
        PlanManifest manifest,
        string relativePath,
        List<string> errors)
    {
        if (manifest.Steps is null)
            return;

        var hasLaunch = false;
        var hasSession = false;

        foreach (var step in manifest.Steps)
        {
            if (!_canonicalByName.TryGetValue(step.Operation, out var descriptor))
                continue;

            if (descriptor.RequiresSession && !hasSession)
            {
                errors.Add(
                    $"{relativePath}: step '{step.Id}' operation '{descriptor.Name}' requires an active session.");
            }

            if (string.Equals(descriptor.Name, "launch", StringComparison.OrdinalIgnoreCase))
            {
                hasLaunch = true;
                if (hasSession)
                {
                    errors.Add($"{relativePath}: duplicate launch without an intervening close or quit.");
                }

                hasSession = true;
                continue;
            }

            if (string.Equals(descriptor.Name, "close", StringComparison.OrdinalIgnoreCase)
                || string.Equals(descriptor.Name, "quit", StringComparison.OrdinalIgnoreCase))
            {
                // Session end only after the RequiresSession check above.
                hasSession = false;
            }
        }

        if (!hasLaunch)
            return;

        var lastStep = manifest.Steps[^1];
        if (!string.Equals(lastStep.Operation, "close", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(lastStep.Operation, "quit", StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"{relativePath}: plans that launch an application must end with close or quit.");
        }

        var onFailure = manifest.OnFailureSteps ?? [];
        if (!onFailure.Any(step =>
                string.Equals(step.Operation, "close", StringComparison.OrdinalIgnoreCase)
                || string.Equals(step.Operation, "quit", StringComparison.OrdinalIgnoreCase)))
        {
            errors.Add(
                $"{relativePath}: plans that launch an application must include close or quit in onFailureSteps.");
        }

        if (hasSession)
        {
            errors.Add($"{relativePath}: plan leaves an active session without close or quit.");
        }
    }
}
