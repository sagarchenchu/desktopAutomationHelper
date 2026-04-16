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
}
