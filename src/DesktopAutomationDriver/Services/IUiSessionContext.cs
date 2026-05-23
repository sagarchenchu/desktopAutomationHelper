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
    /// Closes the active session gracefully: sends WM_CLOSE to tracked windows and
    /// releases automation resources, but does NOT force-kill the process.
    /// </summary>
    void Close();

    /// <summary>
    /// Quits the active session: sends WM_CLOSE to tracked windows, then
    /// force-kills the process tree only if this driver launched it.
    /// Attached sessions are never force-killed unless <paramref name="forceKillAttached"/>
    /// is true.
    /// </summary>
    /// <param name="forceKillAttached">
    /// When true, the process is killed even if this session was attached to an
    /// already-running process.  Use with caution.
    /// </param>
    void Quit(bool forceKillAttached = false);

    /// <summary>
    /// Returns a snapshot of all currently tracked windows for the active session,
    /// or an empty list when no session is active.
    /// </summary>
    IReadOnlyList<TrackedWindowInfo> ListTrackedWindows();
}
