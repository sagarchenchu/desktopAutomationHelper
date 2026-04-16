namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /click/advanced.</summary>
public class ClickAdvancedRequest
{
    /// <summary>UIA Name property of the element to click.</summary>
    public string? Name { get; set; }

    /// <summary>UIA ControlType of the element, e.g. "Button", "Edit".</summary>
    public string? ControlType { get; set; }
}
