namespace DesktopAutomationDriver.Models.Request;

/// <summary>
/// Request model for sending key input to an element.
/// </summary>
public class SendKeysRequest
{
    /// <summary>
    /// Array of characters/strings to type into the element.
    /// Special key codes (e.g. "\uE007" for Enter) are supported.
    /// </summary>
    public string[] Value { get; set; } = Array.Empty<string>();
}
