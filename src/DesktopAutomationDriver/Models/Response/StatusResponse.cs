namespace DesktopAutomationDriver.Models.Response;

/// <summary>
/// Response payload for the /status endpoint.
/// </summary>
public class StatusResponse
{
    /// <summary>
    /// Whether the server is ready to accept new sessions.
    /// </summary>
    public bool Ready { get; set; } = true;

    /// <summary>
    /// Human-readable status message.
    /// </summary>
    public string Message { get; set; } = string.Empty;

    /// <summary>
    /// Driver build / version information.
    /// </summary>
    public BuildInfo Build { get; set; } = new();
}

/// <summary>
/// Build version information about the driver.
/// </summary>
public class BuildInfo
{
    public string Version { get; set; } = "1.0.0";
    public string Revision { get; set; } = string.Empty;
    public string Time { get; set; } = string.Empty;
}
