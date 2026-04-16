namespace DesktopAutomationDriver.Services;

/// <summary>
/// Holds the single "active session" used by the POST /ui endpoint.
/// Callers launch an application once, perform operations, then close.
/// </summary>
public interface IUiSessionContext
{
    /// <summary>
    /// The current active session, or null if no application has been launched.
    /// </summary>
    AutomationSession? ActiveSession { get; }

    /// <summary>
    /// Launches the application at <paramref name="exePath"/> and creates a new
    /// active session. Any previously active session is closed first.
    /// </summary>
    /// <param name="exePath">Full path to the application executable.</param>
    /// <returns>The newly created session.</returns>
    AutomationSession Launch(string exePath);

    /// <summary>
    /// Attaches to an already-running process and creates a new active session.
    /// Any previously active session is closed first.
    /// </summary>
    /// <param name="processId">The PID of the process to attach to.</param>
    /// <returns>The newly created session.</returns>
    AutomationSession Attach(int processId);

    /// <summary>
    /// Closes the active session and terminates the application.
    /// </summary>
    void Close();
}
