namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /listallwindows and POST /maximize.</summary>
public class WindowNameRequest
{
    /// <summary>Partial window title to search for.</summary>
    public string? Window { get; set; }
}
