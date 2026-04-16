namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /switchwindow.</summary>
public class SwitchWindowSimpleRequest
{
    /// <summary>Partial window title to switch focus to.</summary>
    public string? WindowTitle { get; set; }
}
