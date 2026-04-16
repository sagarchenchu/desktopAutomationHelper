namespace DesktopAutomationDriver.Models.Response;

/// <summary>
/// Response payload for a successfully created session.
/// </summary>
public class SessionResponse
{
    /// <summary>
    /// The unique session identifier.
    /// </summary>
    public string SessionId { get; set; } = string.Empty;

    /// <summary>
    /// Capabilities that were granted for this session.
    /// </summary>
    public SessionCapabilities Capabilities { get; set; } = new();
}

/// <summary>
/// Describes the capabilities active for a session.
/// </summary>
public class SessionCapabilities
{
    public string? App { get; set; }
    public string? AppName { get; set; }
    public string UiaType { get; set; } = "UIA3";
    public int ProcessId { get; set; }
}
