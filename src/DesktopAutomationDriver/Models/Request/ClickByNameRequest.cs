namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /click/name and POST /doubleclick/name.</summary>
public class ClickByNameRequest
{
    /// <summary>UIA Name property of the element to click.</summary>
    public string? Name { get; set; }
}
