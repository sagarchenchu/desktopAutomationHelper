using System.Collections.Concurrent;
using System.Diagnostics;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core;
using FlaUI.UIA2;
using FlaUI.UIA3;

// Alias to avoid ambiguity with System.Windows.Forms.Application (added via UseWindowsForms)
using FlaUIApplication = FlaUI.Core.Application;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Thread-safe manager that creates, tracks and disposes automation sessions.
/// </summary>
public class SessionManager : ISessionManager, IDisposable
{
    private readonly ConcurrentDictionary<string, AutomationSession> _sessions = new();
    private readonly ILogger<SessionManager> _logger;
    private bool _disposed;

    public SessionManager(ILogger<SessionManager> logger)
    {
        _logger = logger;
    }

    /// <inheritdoc/>
    public AutomationSession CreateSession(DesiredCapabilities capabilities)
    {
        ArgumentNullException.ThrowIfNull(capabilities);

        AutomationBase automation = capabilities.UiaType?.ToUpperInvariant() == "UIA2"
            ? new UIA2Automation()
            : new UIA3Automation();

        FlaUIApplication application;

        if (!string.IsNullOrWhiteSpace(capabilities.App))
        {
            _logger.LogInformation("Launching application: {App}", capabilities.App);
            var launchInfo = new ProcessStartInfo(capabilities.App)
            {
                UseShellExecute = false
            };
            if (!string.IsNullOrWhiteSpace(capabilities.AppArguments))
                launchInfo.Arguments = capabilities.AppArguments;
            if (!string.IsNullOrWhiteSpace(capabilities.AppWorkingDir))
                launchInfo.WorkingDirectory = capabilities.AppWorkingDir;

            application = FlaUIApplication.Launch(launchInfo);
        }
        else if (!string.IsNullOrWhiteSpace(capabilities.AppName))
        {
            _logger.LogInformation("Attaching to process: {AppName}", capabilities.AppName);
            var process = Process.GetProcessesByName(capabilities.AppName).FirstOrDefault()
                ?? throw new InvalidOperationException(
                    $"No running process found with name '{capabilities.AppName}'.");
            application = FlaUIApplication.Attach(process);
        }
        else
        {
            automation.Dispose();
            throw new ArgumentException(
                "Either 'App' (path to executable) or 'AppName' (process name) must be specified.");
        }

        if (capabilities.LaunchDelay > 0)
        {
            Thread.Sleep(capabilities.LaunchDelay);
        }

        var sessionId = Guid.NewGuid().ToString("N");
        var session = new AutomationSession(
            sessionId,
            application,
            automation,
            capabilities.UiaType ?? "UIA3",
            capabilities.App,
            capabilities.AppName);

        _sessions[sessionId] = session;
        _logger.LogInformation("Session created: {SessionId}", SanitizeId(sessionId));
        return session;
    }

    /// <inheritdoc/>
    public AutomationSession? GetSession(string sessionId) =>
        _sessions.TryGetValue(sessionId, out var session) ? session : null;

    /// <inheritdoc/>
    public void CloseSession(string sessionId)
    {
        if (_sessions.TryRemove(sessionId, out var session))
        {
            _logger.LogInformation("Closing session: {SessionId}", SanitizeId(sessionId));
            session.Dispose();
        }
    }

    /// <inheritdoc/>
    public IReadOnlyList<string> GetAllSessionIds() =>
        _sessions.Keys.ToList();

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        foreach (var session in _sessions.Values)
        {
            try { session.Dispose(); } catch { /* best effort */ }
        }
        _sessions.Clear();
    }

    /// <summary>
    /// Strips control characters from a session ID before including it in a
    /// log message to prevent log-injection attacks.
    /// </summary>
    private static string SanitizeId(string id) =>
        System.Text.RegularExpressions.Regex.Replace(id ?? string.Empty, @"[\r\n\t]", "_");
}
