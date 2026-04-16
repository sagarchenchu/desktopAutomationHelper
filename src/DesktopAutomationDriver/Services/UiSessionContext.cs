using System.Diagnostics;
using FlaUI.UIA3;

// Alias to avoid ambiguity with System.Windows.Forms.Application
using FlaUIApplication = FlaUI.Core.Application;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Thread-safe holder for the single active session used by the POST /ui endpoint.
/// </summary>
public class UiSessionContext : IUiSessionContext, IDisposable
{
    private readonly ILogger<UiSessionContext> _logger;
    private AutomationSession? _activeSession;
    private readonly object _lock = new();
    private bool _disposed;

    public UiSessionContext(ILogger<UiSessionContext> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc/>
    public AutomationSession? ActiveSession
    {
        get { lock (_lock) { return _activeSession; } }
    }

    /// <inheritdoc/>
    public AutomationSession Launch(string exePath)
    {
        lock (_lock)
        {
            // Dispose any existing session before starting a new one.
            if (_activeSession != null)
            {
                _logger.LogInformation("Closing existing UI session before launching new app.");
                _activeSession.Dispose();
                _activeSession = null;
            }

            var automation = new UIA3Automation();
            var launchInfo = new ProcessStartInfo(exePath) { UseShellExecute = false };
            var app = FlaUIApplication.Launch(launchInfo);

            // Brief pause so the application window is ready.
            Thread.Sleep(1000);

            var sessionId = Guid.NewGuid().ToString("N");
            var session = new AutomationSession(
                sessionId, app, automation, "UIA3", exePath, null);

            _activeSession = session;
            _logger.LogInformation("UI session launched for: {ExePath}",
                SanitizePath(exePath));

            // Seed known window handles so windows already open at launch are not
            // treated as "new" the first time GetWindowRoot checks for new windows.
            SeedKnownWindows(session);
            return session;
        }
    }

    /// <inheritdoc/>
    public AutomationSession Attach(int processId)
    {
        lock (_lock)
        {
            if (_activeSession != null)
            {
                _logger.LogInformation(
                    "Closing existing UI session before attaching to process {ProcessId}.", processId);
                _activeSession.Dispose();
                _activeSession = null;
            }

            Process process;
            try
            {
                process = Process.GetProcessById(processId);
            }
            catch (ArgumentException ex)
            {
                throw new InvalidOperationException(
                    $"Cannot attach: no process with ID {processId} is running.", ex);
            }

            var automation = new UIA3Automation();
            var app = FlaUIApplication.Attach(process);

            var sessionId = Guid.NewGuid().ToString("N");
            var session = new AutomationSession(
                sessionId, app, automation, "UIA3", null, SanitizePath(process.ProcessName));

            _activeSession = session;
            _logger.LogInformation("UI session attached to process {ProcessId} ({ProcessName}).",
                processId, SanitizePath(process.ProcessName));

            // Seed known window handles so windows already open at attach are not
            // treated as "new" the first time GetWindowRoot checks for new windows.
            SeedKnownWindows(session);
            return session;
        }
    }

    /// <inheritdoc/>
    public void Close()
    {
        lock (_lock)
        {
            if (_activeSession == null) return;
            _logger.LogInformation("UI session closed.");
            _activeSession.Dispose();
            _activeSession = null;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Close();
    }

    private static string SanitizePath(string? path) =>
        System.Text.RegularExpressions.Regex.Replace(
            path ?? string.Empty, @"[\r\n\t]", "_");

    /// <summary>
    /// Populates the session's set of known window handles from the application's
    /// current top-level windows, so those windows are not treated as "new" by the
    /// auto-follow logic.
    /// </summary>
    private static void SeedKnownWindows(AutomationSession session)
    {
        try
        {
            var windows = session.Application.GetAllTopLevelWindows(session.Automation);
            session.SeedWindowHandles(windows
                .Select(w =>
                {
                    try { return w.Properties.NativeWindowHandle.Value; }
                    catch { return IntPtr.Zero; }
                })
                .Where(h => h != IntPtr.Zero));
        }
        catch { /* best effort – auto-follow will still work, just may pick up existing windows once */ }
    }
}
