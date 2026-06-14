namespace DesktopAutomationDriver.Services.Resolution;

public sealed class ElementSnapshot
{
    public string Name { get; init; } = "";
    public string AutomationId { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string ControlType { get; init; } = "";
    public string FrameworkId { get; init; } = "";
    public string Value { get; init; } = "";
    public string Text { get; init; } = "";
    public string LegacyName { get; init; } = "";
    public string LegacyValue { get; init; } = "";

    public int? Hwnd { get; init; }
    public int? ProcessId { get; init; }
    public int? ControlId { get; init; }

    public bool? IsEnabled { get; init; }
    public bool? IsOffscreen { get; init; }
    public bool? IsVisible { get; init; }

    public object? Rectangle { get; init; }

    public bool HasInvoke { get; init; }
    public bool HasValue { get; init; }
    public bool HasSelectionItem { get; init; }
    public bool HasToggle { get; init; }
    public bool HasScrollItem { get; init; }
    public bool HasExpandCollapse { get; init; }
}

