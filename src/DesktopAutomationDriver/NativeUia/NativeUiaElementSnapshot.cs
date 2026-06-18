namespace DesktopAutomationDriver.NativeUia;

internal sealed class NativeUiaElementSnapshot
{
    public string Name { get; init; } = "";
    public string AutomationId { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string ControlType { get; init; } = "";
    public int ControlTypeId { get; init; }
    public int? ProcessId { get; init; }
    public int? NativeWindowHandle { get; init; }
    public string FrameworkId { get; init; } = "";
    public bool? IsEnabled { get; init; }
    public bool? IsOffscreen { get; init; }
    public object? BoundingRectangle { get; init; }
    public string Value { get; init; } = "";
    public string Text { get; init; } = "";
    public string RuntimeId { get; init; } = "";
    public string MatchText { get; init; } = "";
}
