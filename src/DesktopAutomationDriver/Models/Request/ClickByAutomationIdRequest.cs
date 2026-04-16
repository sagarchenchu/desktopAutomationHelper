namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /click/aid and POST /doubleclick/aid.</summary>
public class ClickByAutomationIdRequest
{
    /// <summary>UIA AutomationId property of the element to click.</summary>
    public string? AutomationId { get; set; }
}
