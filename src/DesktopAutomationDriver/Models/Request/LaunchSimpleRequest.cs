namespace DesktopAutomationDriver.Models.Request;

/// <summary>Request body for POST /launch.</summary>
public class LaunchSimpleRequest
{
    /// <summary>Full path to the executable to launch.</summary>
    public string? ExePath { get; set; }
}
