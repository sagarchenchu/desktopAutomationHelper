namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementSnapshot
{
    public string Name { get; init; } = string.Empty;
    public string AutomationId { get; init; } = string.Empty;
    public string ClassName { get; init; } = string.Empty;
    public string ControlType { get; init; } = string.Empty;
    public string LocalizedControlType { get; init; } = string.Empty;
    public string FrameworkId { get; init; } = string.Empty;
    public string Value { get; init; } = string.Empty;

    public long? Hwnd { get; init; }
    public int? ProcessId { get; init; }
    public string RuntimeId { get; init; } = string.Empty;

    public bool? IsEnabled { get; init; }
    public bool? IsOffscreen { get; init; }
    public bool? HasKeyboardFocus { get; init; }

    public object? Rectangle { get; init; }

    public bool SupportsInvoke { get; init; }
    public bool SupportsValue { get; init; }
    public bool SupportsSelectionItem { get; init; }
    public bool SupportsToggle { get; init; }
    public bool SupportsExpandCollapse { get; init; }
    public bool SupportsScrollItem { get; init; }
}
