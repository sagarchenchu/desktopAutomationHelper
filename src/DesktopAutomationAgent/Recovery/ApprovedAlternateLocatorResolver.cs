using System.Text.Json;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Resolves approved alternate locators for a step without mutating the original plan.
/// </summary>
public sealed class ApprovedAlternateLocatorResolver
{
    public IReadOnlyList<ObjectLocator> GetApprovedAlternates(
        string? objectReference,
        ObjectRepositorySnapshot? snapshot)
    {
        if (string.IsNullOrWhiteSpace(objectReference) || snapshot is null)
            return Array.Empty<ObjectLocator>();

        var parts = objectReference.Split('.', 2);
        if (parts.Length != 2)
            return Array.Empty<ObjectLocator>();

        if (!snapshot.Pages.TryGetValue(parts[0], out var page)
            || page.Elements is null
            || !page.Elements.TryGetValue(parts[1], out var element)
            || element.AlternateLocators is null
            || element.AlternateLocators.Count == 0)
        {
            return Array.Empty<ObjectLocator>();
        }

        var result = new List<ObjectLocator>();
        for (var i = 0; i < element.AlternateLocators.Count; i++)
        {
            var locator = element.AlternateLocators[i];
            var validation = ObjectLocatorValidator.Validate(locator, $"{objectReference}.alternateLocators[{i}]");
            if (validation.IsValid)
                result.Add(locator);
        }

        return result;
    }

    public PlanStep CloneStepWithLocator(PlanStep original, ObjectLocator locator, string locatorKey = "locator")
    {
        ArgumentNullException.ThrowIfNull(original);
        ArgumentNullException.ThrowIfNull(locator);

        var args = original.Arguments is null
            ? new Dictionary<string, JsonElement>(StringComparer.Ordinal)
            : new Dictionary<string, JsonElement>(original.Arguments, StringComparer.Ordinal);

        args[locatorKey] = ObjectLocatorSerializer.ToJsonElement(locator);

        return new PlanStep
        {
            Id = original.Id,
            Operation = original.Operation,
            Arguments = args,
            Assertions = original.Assertions,
            Sensitive = original.Sensitive,
            CaptureResponse = original.CaptureResponse
        };
    }

    public PlanStep? CloneStepWithEquivalentTransition(
        PlanStep original,
        ApprovedEquivalentTransition transition)
    {
        ArgumentNullException.ThrowIfNull(original);
        ArgumentNullException.ThrowIfNull(transition);

        if (string.IsNullOrWhiteSpace(transition.EquivalentAction))
            return null;

        var args = original.Arguments is null
            ? new Dictionary<string, JsonElement>(StringComparer.Ordinal)
            : new Dictionary<string, JsonElement>(original.Arguments, StringComparer.Ordinal);

        if (transition.ReplacementArguments is not null)
        {
            foreach (var pair in transition.ReplacementArguments)
                args[pair.Key] = pair.Value.Clone();
        }

        return new PlanStep
        {
            Id = original.Id,
            Operation = transition.EquivalentAction,
            Arguments = args,
            Assertions = original.Assertions,
            Sensitive = original.Sensitive,
            CaptureResponse = original.CaptureResponse
        };
    }
}
