namespace DesktopAutomationDriver.Services.ElementResolution;

/// <summary>
/// Frozen view of a resolved UIA element for diagnostics and API responses.
/// </summary>
public sealed class ElementSnapshot
{
    public string? Name { get; init; }
    public string? AutomationId { get; init; }
    public string? ClassName { get; init; }
    public string? ControlType { get; init; }
    public string? FrameworkId { get; init; }
    public string? Value { get; init; }

    public int? ProcessId { get; init; }
    public long? Hwnd { get; init; }

    public bool? IsEnabled { get; init; }
    public bool? IsOffscreen { get; init; }

    public object? Rectangle { get; init; }

    public bool HasInvokePattern { get; init; }
    public bool HasValuePattern { get; init; }
    public bool HasSelectionItemPattern { get; init; }
    public bool HasTogglePattern { get; init; }
    public bool HasExpandCollapsePattern { get; init; }
    public bool HasScrollItemPattern { get; init; }

    public static ElementSnapshot FromResolutionSnapshot(Resolution.ElementSnapshot source)
    {
        return new ElementSnapshot
        {
            Name = NullIfEmpty(source.Name),
            AutomationId = NullIfEmpty(source.AutomationId),
            ClassName = NullIfEmpty(source.ClassName),
            ControlType = NullIfEmpty(source.ControlType),
            FrameworkId = NullIfEmpty(source.FrameworkId),
            Value = NullIfEmpty(source.Value),
            ProcessId = source.ProcessId,
            Hwnd = source.Hwnd,
            IsEnabled = source.IsEnabled,
            IsOffscreen = source.IsOffscreen,
            Rectangle = source.Rectangle,
            HasInvokePattern = source.HasInvoke,
            HasValuePattern = source.HasValue,
            HasSelectionItemPattern = source.HasSelectionItem,
            HasTogglePattern = source.HasToggle,
            HasExpandCollapsePattern = source.HasExpandCollapse,
            HasScrollItemPattern = source.HasScrollItem
        };
    }

    private static string? NullIfEmpty(string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
