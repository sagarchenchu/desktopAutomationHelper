using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Manages automation sessions (launch/attach/close applications).
/// </summary>
public interface ISessionManager
{
    /// <summary>
    /// Creates a new automation session, either by launching an application
    /// or attaching to a running process.
    /// </summary>
    /// <param name="capabilities">Capabilities describing the app to automate.</param>
    /// <returns>The created session.</returns>
    AutomationSession CreateSession(DesiredCapabilities capabilities);

    /// <summary>
    /// Retrieves an existing session by its ID.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <returns>The session if found; otherwise null.</returns>
    AutomationSession? GetSession(string sessionId);

    /// <summary>
    /// Closes the session and optionally terminates the application process.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    void CloseSession(string sessionId);

    /// <summary>
    /// Returns all active session IDs.
    /// </summary>
    IReadOnlyList<string> GetAllSessionIds();
}
