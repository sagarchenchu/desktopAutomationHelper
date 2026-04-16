namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Describes a UI element that was identified during recording.
/// </summary>
public class ElementInfo
{
    /// <summary>The element's accessible Name property.</summary>
    public string? Name { get; set; }

    /// <summary>The element's AutomationId property.</summary>
    public string? AutomationId { get; set; }

    /// <summary>The element's ClassName property.</summary>
    public string? ClassName { get; set; }

    /// <summary>The element's control type (e.g. Button, Edit, ListItem).</summary>
    public string? ControlType { get; set; }

    /// <summary>The bounding rectangle of the element on screen.</summary>
    public string? BoundingRectangle { get; set; }

    /// <summary>
    /// A suggested XPath-style selector that can be used to re-locate this
    /// element via the existing element-finding API.
    /// </summary>
    public string? SuggestedXPath { get; set; }

    /// <summary>
    /// Returns a human-readable label for the element using the best available
    /// identifier: Name → AutomationId → ControlType → "(element)".
    /// </summary>
    public static string GetLabel(ElementInfo? info) =>
        !string.IsNullOrEmpty(info?.Name) ? info.Name! :
        !string.IsNullOrEmpty(info?.AutomationId) ? info.AutomationId! :
        !string.IsNullOrEmpty(info?.ControlType) ? info.ControlType! :
        "(element)";
}
