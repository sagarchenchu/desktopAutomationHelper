namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// Live property snapshot of a native UIA element for diagnostics and matching.
/// </summary>
internal sealed class NativeUiaElementSnapshot
{
    public string Name { get; init; } = "";
    public string AutomationId { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string ControlType { get; init; } = "";
    public string FrameworkId { get; init; } = "";
    public string Value { get; init; } = "";
    public string LegacyName { get; init; } = "";
    public string LegacyValue { get; init; } = "";
    public int ProcessId { get; init; }
    public int NativeWindowHandle { get; init; }
    public bool IsEnabled { get; init; }
    public bool IsOffscreen { get; init; }
    public object? BoundingRectangle { get; init; }
    public string RuntimeId { get; init; } = "";
    public List<string> SupportedPatterns { get; init; } = new();

    public object ToDiagnosticObject() => new
    {
        name = Name,
        automationId = AutomationId,
        className = ClassName,
        controlType = ControlType,
        frameworkId = FrameworkId,
        value = Value,
        legacyName = LegacyName,
        legacyValue = LegacyValue,
        processId = ProcessId,
        nativeWindowHandle = NativeWindowHandle,
        isEnabled = IsEnabled,
        isOffscreen = IsOffscreen,
        boundingRectangle = BoundingRectangle,
        runtimeId = RuntimeId,
        supportedPatterns = SupportedPatterns
    };
}
