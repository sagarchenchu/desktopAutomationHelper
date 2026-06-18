using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.ElementResolution;

/// <summary>
/// Serializable search criteria extracted from a <see cref="UiLocator"/> for diagnostics.
/// </summary>
public sealed class ElementSearchCriteria
{
    public string? Name { get; init; }
    public string? AutomationId { get; init; }
    public string? ClassName { get; init; }
    public string? ControlType { get; init; }
    public string? FrameworkId { get; init; }
    public string? Value { get; init; }
    public string? MatchMode { get; init; }
    public string? BestMatch { get; init; }
    public int? ProcessId { get; init; }
    public long? Hwnd { get; init; }
    public int? FoundIndex { get; init; }
    public int? CtrlIndex { get; init; }
    public int? Depth { get; init; }
    public string? SearchScope { get; init; }
    public bool? IncludeOffscreen { get; init; }
    public bool? IncludeDisabled { get; init; }
    public string? NameRegex { get; init; }
    public string? AutomationIdRegex { get; init; }
    public string? ClassNameRegex { get; init; }
    public string? ValueRegex { get; init; }

    public static ElementSearchCriteria FromLocator(UiLocator? locator)
    {
        if (locator == null)
            return new ElementSearchCriteria();

        return new ElementSearchCriteria
        {
            Name = locator.Name,
            AutomationId = locator.AutomationId,
            ClassName = locator.ClassName,
            ControlType = locator.ControlType,
            FrameworkId = locator.FrameworkId,
            Value = locator.Value,
            MatchMode = locator.MatchMode,
            BestMatch = locator.BestMatch,
            ProcessId = locator.ProcessId,
            Hwnd = locator.Hwnd,
            FoundIndex = locator.FoundIndex ?? locator.Index,
            CtrlIndex = locator.CtrlIndex,
            Depth = locator.Depth,
            SearchScope = locator.SearchScope,
            IncludeOffscreen = locator.IncludeOffscreen,
            IncludeDisabled = locator.IncludeDisabled,
            NameRegex = locator.NameRegex,
            AutomationIdRegex = locator.AutomationIdRegex,
            ClassNameRegex = locator.ClassNameRegex,
            ValueRegex = locator.ValueRegex
        };
    }
}
