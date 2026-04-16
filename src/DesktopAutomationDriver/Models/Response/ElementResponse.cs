namespace DesktopAutomationDriver.Models.Response;

/// <summary>
/// Response payload representing a found UI element.
/// </summary>
public class ElementResponse
{
    /// <summary>
    /// The unique element reference identifier used in subsequent element commands.
    /// </summary>
    public string ElementId { get; set; } = string.Empty;
}
