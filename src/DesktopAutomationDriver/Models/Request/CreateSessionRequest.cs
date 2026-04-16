namespace DesktopAutomationDriver.Models.Request;

/// <summary>
/// Request model for creating a new automation session.
/// </summary>
public class CreateSessionRequest
{
    /// <summary>
    /// Desired capabilities for the session.
    /// </summary>
    public DesiredCapabilities DesiredCapabilities { get; set; } = new();
}

/// <summary>
/// Capabilities that define how a session should be created.
/// </summary>
public class DesiredCapabilities
{
    /// <summary>
    /// Full path to the application executable to launch.
    /// Either App or AppName must be specified.
    /// </summary>
    public string? App { get; set; }

    /// <summary>
    /// Process name of an already-running application to attach to.
    /// Either App or AppName must be specified.
    /// </summary>
    public string? AppName { get; set; }

    /// <summary>
    /// Command-line arguments to pass when launching the application.
    /// </summary>
    public string? AppArguments { get; set; }

    /// <summary>
    /// Working directory for the application process.
    /// </summary>
    public string? AppWorkingDir { get; set; }

    /// <summary>
    /// UI Automation type to use: "UIA2" or "UIA3" (default: "UIA3").
    /// </summary>
    public string UiaType { get; set; } = "UIA3";

    /// <summary>
    /// Milliseconds to wait for application to launch before starting automation.
    /// </summary>
    public int LaunchDelay { get; set; } = 1000;
}
