namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>
/// Driver-side mirror of Agent <c>ObjectLocatorValidator</c> identity rules for page candidates.
/// Keeps Assistive export aligned with Phase 3 runtime validation without referencing Agent.
/// </summary>
public static class Phase3LocatorContract
{
    public static bool IsValid(IReadOnlyDictionary<string, object?> locator, out string error)
    {
        error = string.Empty;
        if (locator is null || locator.Count == 0)
        {
            error = "locator must include at least one identifying field.";
            return false;
        }

        var automationId = GetString(locator, "automationId");
        var name = GetString(locator, "name");
        var className = GetString(locator, "className");
        var controlType = GetString(locator, "controlType");

        if (IsBlank(automationId) || IsBlank(name) || IsBlank(className) || IsBlank(controlType))
        {
            error = "locator string fields must not be blank or whitespace.";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(controlType)
            && !Phase3KnownControlTypes.IsKnown(controlType))
        {
            error = $"controlType '{controlType}' is not a recognized UIA control type.";
            return false;
        }

        var hasAutomationId = !string.IsNullOrWhiteSpace(automationId);
        var hasName = !string.IsNullOrWhiteSpace(name);
        var hasClassName = !string.IsNullOrWhiteSpace(className);
        var hasControlType = !string.IsNullOrWhiteSpace(controlType);

        if (!hasAutomationId && !hasName && !hasClassName && !hasControlType)
        {
            error = "locator must include at least one identifying field.";
            return false;
        }

        if (hasAutomationId)
            return true;

        if (hasName && !hasControlType)
        {
            error = "name requires controlType when automationId is absent.";
            return false;
        }

        if (hasClassName && !hasControlType)
        {
            error = "className requires controlType when automationId is absent.";
            return false;
        }

        if (hasControlType && !hasName && !hasClassName)
        {
            error = "controlType alone is not sufficient without automationId.";
            return false;
        }

        if (!hasName && !hasClassName)
        {
            error = "without automationId, locator must include name+controlType or className+controlType.";
            return false;
        }

        return true;
    }

    private static string? GetString(IReadOnlyDictionary<string, object?> locator, string key) =>
        locator.TryGetValue(key, out var value) ? value as string : null;

    private static bool IsBlank(string? value) => value is not null && string.IsNullOrWhiteSpace(value);
}
