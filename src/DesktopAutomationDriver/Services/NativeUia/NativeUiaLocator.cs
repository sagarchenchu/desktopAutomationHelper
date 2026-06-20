namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// pywinauto-style locator criteria for native UIA element search.
/// </summary>
internal sealed class NativeUiaLocator
{
    public string? Name { get; init; }
    public string? AutomationId { get; init; }
    public string? ClassName { get; init; }
    public string? ControlType { get; init; }
    public string? Value { get; init; }
    public long? Hwnd { get; init; }
    public int? ProcessId { get; init; }
    public bool EnforceProcessIdMatch { get; init; }
    public int? FoundIndex { get; init; }
    public string MatchMode { get; init; } = "exact";

    public NativeUiaLocator AsComboBoxLocator() =>
        new()
        {
            Name = Name,
            AutomationId = AutomationId,
            ClassName = ClassName,
            ControlType = "ComboBox",
            Value = Value,
            Hwnd = Hwnd,
            ProcessId = ProcessId,
            EnforceProcessIdMatch = EnforceProcessIdMatch,
            FoundIndex = FoundIndex,
            MatchMode = MatchMode
        };

    public NativeUiaLocator AutomationIdOnly() =>
        new()
        {
            AutomationId = AutomationId,
            ProcessId = ProcessId,
            EnforceProcessIdMatch = EnforceProcessIdMatch,
            Hwnd = Hwnd,
            FoundIndex = FoundIndex,
            MatchMode = MatchMode
        };

    public NativeUiaLocator NameOnly() =>
        new()
        {
            Name = Name,
            ProcessId = ProcessId,
            EnforceProcessIdMatch = EnforceProcessIdMatch,
            Hwnd = Hwnd,
            FoundIndex = FoundIndex,
            MatchMode = MatchMode
        };

    public NativeUiaLocator ControlTypeOnly() =>
        new()
        {
            ControlType = ControlType,
            ProcessId = ProcessId,
            EnforceProcessIdMatch = EnforceProcessIdMatch,
            Hwnd = Hwnd,
            FoundIndex = FoundIndex,
            MatchMode = MatchMode
        };

    public NativeUiaLocator WithoutControlType() =>
        new()
        {
            Name = Name,
            AutomationId = AutomationId,
            ClassName = ClassName,
            Value = Value,
            Hwnd = Hwnd,
            ProcessId = ProcessId,
            EnforceProcessIdMatch = EnforceProcessIdMatch,
            FoundIndex = FoundIndex,
            MatchMode = MatchMode
        };
}
