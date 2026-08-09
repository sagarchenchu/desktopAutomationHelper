using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class PlanObjectReferenceExpander
{
    private static readonly Regex ReferencePattern = new(
        @"^[a-z][a-z0-9-]{0,63}\.[a-z][a-z0-9-]{0,63}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly HashSet<string> AllowedLocatorKeys = new(StringComparer.Ordinal)
    {
        "locator",
        "locator2",
        "parentLocator",
        "containerLocator"
    };

    private readonly ObjectReferenceResolver _resolver;

    public PlanObjectReferenceExpander(ObjectReferenceResolver resolver)
    {
        _resolver = resolver;
    }

    public PlanObjectReferenceExpansionResult Expand(
        PlanManifest plan,
        ObjectRepositorySnapshot snapshot,
        string planPath)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(snapshot);

        var errors = new List<string>();
        var warnings = new List<string>();
        var resolvedRefs = new HashSet<string>(StringComparer.Ordinal);

        var hasObjectRefs = false;
        ValidateAndCollect(plan, planPath, errors, ref hasObjectRefs);

        if (errors.Count > 0)
        {
            return BuildResult(errors, warnings, snapshot, resolvedRefs);
        }

        if (!hasObjectRefs)
        {
            return BuildResult(errors, warnings, snapshot, resolvedRefs);
        }

        if (string.IsNullOrWhiteSpace(plan.ObjectRepository))
        {
            errors.Add($"{planPath}: plans containing $objectRef markers must declare objectRepository.");
            return BuildResult(errors, warnings, snapshot, resolvedRefs);
        }

        ExpandSteps(plan.Steps, planPath, "steps", snapshot, errors, warnings, resolvedRefs);
        ExpandSteps(plan.OnFailureSteps, planPath, "onFailureSteps", snapshot, errors, warnings, resolvedRefs);

        return BuildResult(errors, warnings, snapshot, resolvedRefs);
    }

    private void ValidateAndCollect(
        PlanManifest plan,
        string planPath,
        List<string> errors,
        ref bool hasObjectRefs)
    {
        ValidateStepList(plan.Steps, planPath, "steps", errors, ref hasObjectRefs);
        ValidateStepList(plan.OnFailureSteps, planPath, "onFailureSteps", errors, ref hasObjectRefs);
    }

    private static void ValidateStepList(
        List<PlanStep>? steps,
        string planPath,
        string phase,
        List<string> errors,
        ref bool hasObjectRefs)
    {
        if (steps is null)
            return;

        for (var i = 0; i < steps.Count; i++)
        {
            var step = steps[i];
            var location = $"{planPath}: {phase}[{i}] id='{step.Id}'";
            if (step.Arguments is null)
                continue;

            foreach (var (key, value) in step.Arguments)
            {
                if (AllowedLocatorKeys.Contains(key))
                {
                    if (TryGetObjectRef(value, out var reference))
                    {
                        hasObjectRefs = true;
                        ValidateReferenceFormat(reference, $"{location}.arguments.{key}", errors);
                    }

                    continue;
                }

                ScanForForbiddenMarkers(value, $"{location}.arguments.{key}", errors, ref hasObjectRefs);
            }
        }
    }

    private void ExpandSteps(
        List<PlanStep>? steps,
        string planPath,
        string phase,
        ObjectRepositorySnapshot snapshot,
        List<string> errors,
        List<string> warnings,
        HashSet<string> resolvedRefs)
    {
        if (steps is null)
            return;

        for (var i = 0; i < steps.Count; i++)
        {
            var step = steps[i];
            var location = $"{planPath}: {phase}[{i}] id='{step.Id}'";
            if (step.Arguments is null)
                continue;

            var updated = new Dictionary<string, JsonElement>(step.Arguments, StringComparer.Ordinal);
            var changed = false;

            foreach (var key in AllowedLocatorKeys)
            {
                if (!updated.TryGetValue(key, out var value))
                    continue;

                if (!TryGetObjectRef(value, out var reference))
                    continue;

                var resolution = _resolver.Resolve(snapshot, reference);
                warnings.AddRange(resolution.Warnings);
                if (!resolution.IsResolved || resolution.Locator is null)
                {
                    errors.AddRange(resolution.Errors.Count > 0
                        ? resolution.Errors
                        : [$"{location}.arguments.{key}: failed to resolve '{reference}'."]);
                    continue;
                }

                updated[key] = ObjectLocatorSerializer.ToJsonElement(resolution.Locator);
                resolvedRefs.Add(reference);
                changed = true;
            }

            if (changed)
                step.Arguments = updated;
        }
    }

    private static void ScanForForbiddenMarkers(
        JsonElement element,
        string location,
        List<string> errors,
        ref bool hasObjectRefs)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                if (element.TryGetProperty("$objectRef", out var objectRef))
                {
                    hasObjectRefs = true;
                    if (element.EnumerateObject().Count() != 1)
                    {
                        errors.Add($"{location}: $objectRef marker must be an object with exactly one property.");
                    }
                    else
                    {
                        errors.Add($"{location}: $objectRef is only allowed in locator, locator2, parentLocator, or containerLocator.");
                    }

                    return;
                }

                foreach (var property in element.EnumerateObject())
                {
                    ScanForForbiddenMarkers(property.Value, $"{location}.{property.Name}", errors, ref hasObjectRefs);
                }

                break;

            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in element.EnumerateArray())
                {
                    ScanForForbiddenMarkers(item, $"{location}[{index}]", errors, ref hasObjectRefs);
                    index++;
                }

                break;
        }
    }

    private static bool TryGetObjectRef(JsonElement element, out string reference)
    {
        reference = string.Empty;
        if (element.ValueKind != JsonValueKind.Object)
            return false;

        var properties = element.EnumerateObject().ToList();
        if (properties.Count != 1 || !properties[0].NameEquals("$objectRef"))
            return false;

        if (properties[0].Value.ValueKind != JsonValueKind.String)
            return false;

        reference = properties[0].Value.GetString() ?? string.Empty;
        return true;
    }

    private static void ValidateReferenceFormat(string reference, string location, List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(reference) || !ReferencePattern.IsMatch(reference))
        {
            errors.Add($"{location}: $objectRef '{reference}' must match pageId.elementId.");
        }
    }

    private static PlanObjectReferenceExpansionResult BuildResult(
        List<string> errors,
        List<string> warnings,
        ObjectRepositorySnapshot snapshot,
        HashSet<string> resolvedRefs) =>
        new()
        {
            Errors = errors,
            Warnings = warnings,
            RepositoryPath = snapshot.RepositoryPath,
            RepositoryId = snapshot.Manifest.RepositoryId,
            RepositorySha256 = snapshot.AggregateSha256,
            ResolvedObjectReferences = resolvedRefs.OrderBy(static r => r, StringComparer.Ordinal).ToArray()
        };
}
