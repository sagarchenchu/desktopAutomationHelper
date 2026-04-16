using FlaUI.Core;
using FlaUI.Core.AutomationElements;

// Alias to avoid ambiguity with System.Windows.Forms.Application (added via UseWindowsForms)
using FlaUIApplication = FlaUI.Core.Application;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Represents an active automation session, holding the FlaUI application and
/// automation backend references plus an in-memory element cache.
/// </summary>
public class AutomationSession : IDisposable
{
    private readonly Dictionary<string, AutomationElement> _elementCache = new();
    private bool _disposed;

    public AutomationSession(
        string sessionId,
        FlaUIApplication application,
        AutomationBase automation,
        string uiaType,
        string? appPath,
        string? appName)
    {
        SessionId = sessionId;
        Application = application;
        Automation = automation;
        UiaType = uiaType;
        AppPath = appPath;
        AppName = appName;
        CreatedAt = DateTimeOffset.UtcNow;
    }

    /// <summary>Unique identifier for this session.</summary>
    public string SessionId { get; }

    /// <summary>FlaUI Application wrapper.</summary>
    public FlaUIApplication Application { get; }

    /// <summary>FlaUI UIA2 or UIA3 automation backend.</summary>
    public AutomationBase Automation { get; }

    /// <summary>UI Automation type used ("UIA2" or "UIA3").</summary>
    public string UiaType { get; }

    /// <summary>Path of the launched application, if any.</summary>
    public string? AppPath { get; }

    /// <summary>Process name of an attached application, if any.</summary>
    public string? AppName { get; }

    /// <summary>When this session was created.</summary>
    public DateTimeOffset CreatedAt { get; }

    /// <summary>
    /// The window element that is currently active for the /ui endpoint.
    /// When null, operations fall back to the application's main window.
    /// Updated by the "switchwindow" operation or auto-follow detection.
    /// </summary>
    public AutomationElement? ActiveWindow { get; set; }

    /// <summary>
    /// When true (the default), element operations automatically switch focus to any
    /// new top-level window that opens within the application after the session started,
    /// without requiring an explicit "switchwindow" call.
    /// </summary>
    public bool AutoFollowNewWindows { get; set; } = true;

    // Tracks window handles seen so far; new ones trigger auto-follow.
    private readonly HashSet<IntPtr> _seenWindowHandles = new();
    private readonly object _windowLock = new();

    /// <summary>
    /// Seeds the set of known window handles with the initial application windows,
    /// so that windows already open at launch are not treated as "new" later.
    /// </summary>
    internal void SeedWindowHandles(IEnumerable<IntPtr> handles)
    {
        lock (_windowLock)
            foreach (var h in handles)
                _seenWindowHandles.Add(h);
    }

    /// <summary>
    /// Inspects <paramref name="currentWindows"/>, registers their handles as known,
    /// and returns the first window whose handle was not previously seen.
    /// Returns null when all windows were already known.
    /// </summary>
    internal AutomationElement? ClaimFirstNewWindow(AutomationElement[] currentWindows)
    {
        lock (_windowLock)
        {
            AutomationElement? newest = null;
            foreach (var w in currentWindows)
            {
                IntPtr h;
                try { h = w.Properties.NativeWindowHandle.Value; }
                catch { continue; }

                if (h != IntPtr.Zero && _seenWindowHandles.Add(h))
                    newest = w;
            }
            return newest;
        }
    }

    /// <summary>
    /// Caches a UI element and returns a stable string ID for it.
    /// </summary>
    public string CacheElement(AutomationElement element)
    {
        var id = Guid.NewGuid().ToString("N");
        _elementCache[id] = element;
        return id;
    }

    /// <summary>
    /// Retrieves a cached UI element by its ID.
    /// </summary>
    public AutomationElement? GetCachedElement(string elementId) =>
        _elementCache.TryGetValue(elementId, out var el) ? el : null;

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _elementCache.Clear();
        CloseAllApplicationWindows();
        try { Automation.Dispose(); } catch { /* best effort */ }
        try { Application.Dispose(); } catch { /* best effort */ }
    }

    /// <summary>
    /// Sends a graceful close message to every top-level window of the application,
    /// then kills the process (including its entire child-process tree) as a safety
    /// net for windows that did not respond to the close message.
    /// </summary>
    private void CloseAllApplicationWindows()
    {
        // Step 1: graceful close via the Window UIA pattern.
        try
        {
            var windows = Application.GetAllTopLevelWindows(Automation);
            foreach (var w in windows)
            {
                try { w.Patterns.Window.PatternOrDefault?.Close(); }
                catch { /* best effort */ }
            }
        }
        catch { /* best effort */ }

        // Step 2: kill the process and its entire child-process tree so that any
        // windows not handled by the graceful close are forcefully terminated.
        try
        {
            var proc = System.Diagnostics.Process.GetProcessById(Application.ProcessId);
            if (!proc.HasExited)
                proc.Kill(entireProcessTree: true);
        }
        catch { /* best effort */ }

        // Step 3: fallback via FlaUI's own kill helper (handles the case where the
        // process ID look-up above failed but the FlaUI handle is still valid).
        try { Application.Kill(); } catch { /* best effort */ }
    }
}
