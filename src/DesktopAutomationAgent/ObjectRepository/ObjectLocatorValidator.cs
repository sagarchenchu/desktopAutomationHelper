namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectLocatorValidationResult
{
    public bool IsValid => Errors.Count == 0;

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
}

public static class ObjectLocatorValidator
{
  private static readonly HashSet<string> AllowedMatchModes = new(StringComparer.Ordinal)
    {
        "exact",
        "contains",
        "startswith"
    };

    private static readonly HashSet<string> VolatileExtensionKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "processId",
        "pid",
        "hwnd",
        "handle",
        "runtimeId",
        "boundingRectangle",
        "left",
        "top",
        "right",
        "bottom",
        "width",
        "height",
        "nearX",
        "nearY",
        "coordinates"
    };

    public static ObjectLocatorValidationResult Validate(ObjectLocator? locator, string location)
    {
        if (locator is null)
        {
            return new ObjectLocatorValidationResult
            {
                Errors = [$"{location}: locator is required."]
            };
        }

        var errors = new List<string>();
        var warnings = new List<string>();

        if (locator.ExtensionData is { Count: > 0 })
        {
            foreach (var key in locator.ExtensionData.Keys)
            {
                if (VolatileExtensionKeys.Contains(key))
                {
                    errors.Add($"{location}: volatile locator property '{key}' is not allowed.");
                }
                else
                {
                    errors.Add($"{location}: unknown locator property '{key}'.");
                }
            }
        }

        RejectBlankString(locator.AutomationId, "automationId", location, errors);
        RejectBlankString(locator.Name, "name", location, errors);
        RejectBlankString(locator.ClassName, "className", location, errors);
        RejectBlankString(locator.ControlType, "controlType", location, errors);
        RejectBlankString(locator.MatchMode, "matchMode", location, errors);

        if (!string.IsNullOrWhiteSpace(locator.ControlType)
            && !LocatorMatchNormalizer.IsKnownControlType(locator.ControlType))
        {
            errors.Add(
                $"{location}: controlType '{locator.ControlType}' is not a recognized UIA control type.");
        }

        if (!string.IsNullOrWhiteSpace(locator.MatchMode)
            && !AllowedMatchModes.Contains(locator.MatchMode))
        {
            errors.Add(
                $"{location}: matchMode '{locator.MatchMode}' is not supported; expected exact, contains, or startswith.");
        }

        if (locator.FoundIndex is < 0)
        {
            errors.Add($"{location}: foundIndex must be greater than or equal to 0.");
        }
        else if (locator.FoundIndex is not null)
        {
            warnings.Add($"{location}: foundIndex makes the locator fragile when UI layout changes.");
        }

        var hasAutomationId = HasValue(locator.AutomationId);
        var hasName = HasValue(locator.Name);
        var hasClassName = HasValue(locator.ClassName);
        var hasControlType = HasValue(locator.ControlType);

        if (!hasAutomationId && !hasName && !hasClassName && !hasControlType
            && string.IsNullOrWhiteSpace(locator.MatchMode)
            && locator.FoundIndex is null)
        {
            errors.Add($"{location}: locator must include at least one identifying field.");
            return new ObjectLocatorValidationResult { Errors = errors, Warnings = warnings };
        }

        if (hasAutomationId)
        {
            return new ObjectLocatorValidationResult { Errors = errors, Warnings = warnings };
        }

        if (hasName && !hasControlType)
        {
            errors.Add($"{location}: name requires controlType when automationId is absent.");
        }

        if (hasClassName && !hasControlType)
        {
            errors.Add($"{location}: className requires controlType when automationId is absent.");
        }

        if (hasControlType && !hasName && !hasClassName)
        {
            errors.Add($"{location}: controlType alone is not sufficient without automationId.");
        }

        if (!hasName && !hasClassName)
        {
            errors.Add(
                $"{location}: without automationId, locator must include name+controlType or className+controlType.");
        }

        return new ObjectLocatorValidationResult { Errors = errors, Warnings = warnings };
    }

    private static bool HasValue(string? value) => !string.IsNullOrWhiteSpace(value);

    private static void RejectBlankString(
        string? value,
        string propertyName,
        string location,
        List<string> errors)
    {
        if (value is not null && string.IsNullOrWhiteSpace(value))
        {
            errors.Add($"{location}: {propertyName} must not be blank or whitespace.");
        }
    }
}
