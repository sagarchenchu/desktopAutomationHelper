using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using System.Runtime.InteropServices;

// Alias to avoid ambiguity with System.Windows.Forms.Application (added via UseWindowsForms)
using FlaUIApplication = FlaUI.Core.Application;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Holds information about a top-level window that was opened within or by the
/// automation session.  Used for reliable cleanup when the session is closed.
/// </summary>
public sealed class TrackedWindowInfo
{
    public IntPtr Hwnd { get; init; }
    public int ProcessId { get; init; }
    public string Title { get; init; } = string.Empty;
    public string ClassName { get; init; } = string.Empty;
    public DateTime FirstSeenUtc { get; init; } = DateTime.UtcNow;
    public DateTime LastSeenUtc { get; set; } = DateTime.UtcNow;
    public bool IsMainWindow { get; init; }
}

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
    /// Process ID of the application launched by this session.
    /// Null for sessions that attached to an already-running process.
    /// </summary>
    public int? LaunchedProcessId { get; set; }

    /// <summary>
    /// When the session application was launched or attached.
    /// </summary>
    public DateTime LaunchUtc { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// True when this driver started the process (via <see cref="IUiSessionContext.Launch"/>).
    /// False for sessions that attached to an already-running process.
    /// Only launched sessions may be force-killed.
    /// </summary>
    public bool WasLaunchedByDriver => LaunchedProcessId.HasValue;

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

    // Ownership tracking: all windows opened by or from this session.
    private readonly Dictionary<IntPtr, TrackedWindowInfo> _trackedWindows = new();

    /// <summary>
    /// Timestamp of the last full desktop-descendant scan in <c>GetWindowRoot</c>.
    /// Used to throttle the expensive scan so it runs at most once per
    /// <see cref="DesktopScanThrottle"/> interval.
    /// </summary>
    internal DateTime LastDesktopScan { get; set; } = DateTime.MinValue;

    /// <summary>
    /// Minimum interval between consecutive full desktop-descendant scans.
    /// The scan covers all Window-type descendants across all processes and is
    /// throttled to avoid overhead during rapid successive operations.
    /// </summary>
    internal static readonly TimeSpan DesktopScanThrottle = TimeSpan.FromSeconds(2);

    /// <summary>
    /// Brief pause after broadcasting WM_CLOSE to tracked windows so the messages
    /// can be processed before the UIA Window-pattern close attempt begins.
    /// </summary>
    private const int WmCloseProcessingDelayMs = 200;

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
    /// Registers a window handle for ownership tracking.  The handle is also seeded
    /// into the set of known handles so auto-follow does not treat it as "new".
    /// Calling again for an already-tracked HWND refreshes <see cref="TrackedWindowInfo.LastSeenUtc"/>.
    /// </summary>
    internal void TrackWindow(
        IntPtr hwnd, int processId, string title, string className, bool isMainWindow)
    {
        if (hwnd == IntPtr.Zero)
            return;

        lock (_windowLock)
        {
            _seenWindowHandles.Add(hwnd);

            if (_trackedWindows.TryGetValue(hwnd, out var existing))
            {
                existing.LastSeenUtc = DateTime.UtcNow;
            }
            else
            {
                _trackedWindows[hwnd] = new TrackedWindowInfo
                {
                    Hwnd        = hwnd,
                    ProcessId   = processId,
                    Title       = title,
                    ClassName   = className,
                    IsMainWindow = isMainWindow,
                };
            }
        }
    }

    /// <summary>
    /// Returns a snapshot of all currently tracked windows.
    /// </summary>
    internal IReadOnlyList<TrackedWindowInfo> GetTrackedWindows()
    {
        lock (_windowLock)
            return _trackedWindows.Values.ToList();
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
    /// Releases the FlaUI automation and application references without closing any
    /// windows or killing the process.  Called by the graceful-close path where the
    /// caller has already sent WM_CLOSE to tracked windows.
    /// </summary>
    internal void ReleaseAutomationResources()
    {
        if (_disposed) return;
        _disposed = true;
        _elementCache.Clear();
        try { Automation.Dispose(); } catch { /* best effort */ }
        try { Application.Dispose(); } catch { /* best effort */ }
    }

    /// <summary>
    /// Sends WM_CLOSE to every tracked window and then attempts a graceful UIA
    /// close for remaining top-level windows.  Does NOT kill the process.
    /// </summary>
    internal void CloseWindowsGracefully()
    {
        // Step 1: send WM_CLOSE to all tracked HWNDs (popup/dialog windows that
        // may no longer appear in GetAllTopLevelWindows).
        try
        {
            foreach (var info in GetTrackedWindows())
            {
                if (info.Hwnd != IntPtr.Zero)
                    PostMessage(info.Hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
        }
        catch { /* best effort */ }

        Thread.Sleep(WmCloseProcessingDelayMs); // allow WM_CLOSE messages to be processed

        // Step 2: graceful close via the Window UIA pattern for remaining windows.
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
    }

    /// <summary>
    /// Sends a graceful close message to every top-level window of the application,
    /// then kills the process tree — but only when this driver launched the process
    /// (<see cref="WasLaunchedByDriver"/> is true).  Attached sessions are never
    /// force-killed; only their windows receive WM_CLOSE.
    /// </summary>
    private void CloseAllApplicationWindows()
    {
        // Graceful close via WM_CLOSE and UIA Window.Close().
        CloseWindowsGracefully();

        // Brief additional wait before the force-kill so the application can process
        // WM_CLOSE and dismiss any "save changes?" dialogs.
        Thread.Sleep(500);

        // Force-kill only when this driver owns the process.
        if (!WasLaunchedByDriver)
            return;

        // NOTE: HasExited guard intentionally absent — Kill(entireProcessTree:true) uses
        // CreateToolhelp32Snapshot and can kill orphaned children even when the parent
        // has already exited; calling Kill() on an exited parent is a safe no-op.
        System.Diagnostics.Process? proc = null;
        try
        {
            proc = System.Diagnostics.Process.GetProcessById(Application.ProcessId);
            proc.Kill(entireProcessTree: true);
        }
        catch { /* best effort */ }

        // Fallback via FlaUI's own kill helper.
        try { Application.Kill(); } catch { /* best effort */ }

        // Wait for the process to fully exit.
        try { proc?.WaitForExit(3000); } catch { /* best effort */ }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private const uint WM_CLOSE = 0x0010;
}
