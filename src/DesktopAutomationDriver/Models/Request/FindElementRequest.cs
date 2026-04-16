namespace DesktopAutomationDriver.Models.Request;

/// <summary>
/// Request model for finding an element within a session or parent element.
/// </summary>
public class FindElementRequest
{
    /// <summary>
    /// The element location strategy to use.
    /// Supported values: "automation id", "id", "name", "class name", "tag name",
    /// "link text", "partial link text", "xpath"
    /// </summary>
    public string Using { get; set; } = string.Empty;

    /// <summary>
    /// The value to search for using the given strategy.
    /// </summary>
    public string Value { get; set; } = string.Empty;
}
