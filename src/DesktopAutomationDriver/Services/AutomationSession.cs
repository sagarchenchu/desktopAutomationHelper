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
        try { Automation.Dispose(); } catch { /* best effort */ }
        try { Application.Dispose(); } catch { /* best effort */ }
    }
}
