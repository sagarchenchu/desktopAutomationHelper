namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /close.</summary>
public class CloseWindowRequest
{
    /// <summary>Partial window title of the window to close.</summary>
    public string? App { get; set; }
}
