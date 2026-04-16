namespace DesktopAutomationDriver.Models.Request;

/// <summary>
/// Identifies a UI element for a /ui endpoint operation.
/// Multiple properties are combined with AND logic.
/// </summary>
public class UiLocator
{
    /// <summary>UIA Name property (element label).</summary>
    public string? Name { get; set; }

    /// <summary>UIA AutomationId property.</summary>
    public string? AutomationId { get; set; }

    /// <summary>UIA ClassName property.</summary>
    public string? ClassName { get; set; }

    /// <summary>
    /// UIA control type string, e.g. "Button", "Edit", "ComboBox", "CheckBox", "Text".
    /// </summary>
    public string? ControlType { get; set; }
}
