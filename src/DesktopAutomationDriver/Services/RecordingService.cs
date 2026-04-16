using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationDriver.Models.Recording;
using FlaUI.UIA3;

// Alias to avoid ambiguity with FlaUI.Core.Application (both in scope via implicit usings)
using WinForms = System.Windows.Forms;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Singleton service that manages the window-activity recording session.
/// It owns the recorded-actions log and JSON export logic.
/// The WinForms overlay window (and all UIA automation) lives on a dedicated STA thread.
/// </summary>
public sealed class RecordingService : IRecordingService, IDisposable
{
    private readonly ILogger<RecordingService> _logger;

    private readonly List<RecordedAction> _actions = [];
    private readonly object _lock = new();

    private volatile bool _isActive;
    private volatile RecordingMode _currentMode = RecordingMode.None;

    private DateTimeOffset _startedAt;
    private DateTimeOffset? _stoppedAt;
    private string? _exportFilePath;

    private Thread? _overlayThread;
    private RecordingOverlayWindow? _overlayWindow;

    // Used by the overlay to run FromPoint on the STA thread
    private UIA3Automation? _automation;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    public RecordingService(ILogger<RecordingService> logger)
    {
        _logger = logger;
    }

    // ── IRecordingService ────────────────────────────────────────────────────

    public bool IsActive => _isActive;

    public RecordingMode CurrentMode => _currentMode;

    public DateTimeOffset? StartedAt => _isActive ? _startedAt : null;

    public string StartRecording()
    {
        if (_isActive)
            return "Recording is already active.";

        lock (_lock)
        {
            _actions.Clear();
            _exportFilePath = null;
            _stoppedAt = null;
            _currentMode = RecordingMode.None;
            _startedAt = DateTimeOffset.UtcNow;
            _isActive = true;
        }

        // Start the overlay on a dedicated STA thread (required for WinForms + COM/UIA).
        // The thread is IsBackground = true, so the runtime can exit when the ASP.NET host
        // shuts down.  If recording is still active at that point the in-process JSON export
        // will not run.  Callers should call StopRecording() (or the user should press Ctrl+S)
        // before shutting down the host to ensure the export file is written.
        _overlayThread = new Thread(() =>
        {
            WinForms.Application.EnableVisualStyles();
            WinForms.Application.SetCompatibleTextRenderingDefault(false);

            // Create UIA automation on the STA thread so COM apartment is correct
            try
            {
                _automation = new UIA3Automation();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not create UIA3Automation on overlay thread");
            }

            _overlayWindow = new RecordingOverlayWindow(this, _logger);
            WinForms.Application.Run(_overlayWindow);

            // Clean up after the form closes
            _automation?.Dispose();
            _automation = null;
            _overlayWindow = null;
        })
        {
            IsBackground = true,
            Name = "RecordingOverlay-STA"
        };
        _overlayThread.SetApartmentState(ApartmentState.STA);
        _overlayThread.Start();

        _logger.LogInformation("Recording session started at {Time}", _startedAt);
        return string.Empty; // no error
    }

    public RecordingExport StopRecording()
    {
        // Close the overlay if it is still open (thread-safe)
        CloseOverlayIfOpen();

        // If already stopped by Ctrl+S, just return current state
        return BuildExport();
    }

    public RecordingExport GetCurrentState() => BuildExport();

    public void SetMode(RecordingMode mode)
    {
        _currentMode = mode;
        _logger.LogInformation("Recording mode changed to {Mode}", mode);
    }

    public void AddAction(RecordedAction action)
    {
        action.Timestamp = DateTimeOffset.UtcNow;
        action.Mode = _currentMode;

        // Generate a fallback description if the caller did not supply one
        if (string.IsNullOrEmpty(action.Description))
        {
            var elementLabel = ElementInfo.GetLabel(action.Element);
            action.Description = action.QueryResult.HasValue
                ? $"{action.ActionType} check on {elementLabel}: {action.QueryResult}"
                : $"{action.ActionType} on {elementLabel}";
        }

        lock (_lock) { _actions.Add(action); }
        _logger.LogDebug("Recorded action: {Type} on [{Element}]",
            action.ActionType, action.Element?.ControlType ?? "?");
    }

    public ElementInfo? GetElementAtPoint(System.Drawing.Point point)
    {
        if (_automation == null) return null;
        try
        {
            var element = _automation.FromPoint(point);
            return element == null ? null : RecordingOverlayWindow.BuildElementInfo(element);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "GetElementAtPoint failed for {Point}", point);
            return null;
        }
    }

    public (ElementInfo? info, IReadOnlyList<ElementInfo> children) GetElementWithChildrenAtPoint(
        System.Drawing.Point point)
    {
        if (_automation == null) return (null, Array.Empty<ElementInfo>());
        try
        {
            var element = _automation.FromPoint(point);
            if (element == null) return (null, Array.Empty<ElementInfo>());

            var info = RecordingOverlayWindow.BuildElementInfo(element);
            var childInfos = element.FindAllChildren()
                .Select(RecordingOverlayWindow.BuildElementInfo)
                .ToArray();
            return (info, childInfos);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "GetElementWithChildrenAtPoint failed for {Point}", point);
            return (null, Array.Empty<ElementInfo>());
        }
    }

    public void OnOverlayClosed()
    {
        if (!_isActive) return;

        _isActive = false;
        _stoppedAt = DateTimeOffset.UtcNow;

        // Export JSON
        try
        {
            ExportJson();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to export recording JSON");
        }

        _logger.LogInformation(
            "Recording stopped at {Time}. {Count} action(s) recorded. Export: {Path}",
            _stoppedAt, _actions.Count, _exportFilePath ?? "(none)");
    }

    public void Dispose()
    {
        CloseOverlayIfOpen();
        _automation?.Dispose();
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private void CloseOverlayIfOpen()
    {
        var window = _overlayWindow;
        if (window == null) return;

        try
        {
            if (window.IsHandleCreated && !window.IsDisposed)
                window.Invoke(new Action(window.Close));
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Error closing overlay window");
        }
    }

    private void ExportJson()
    {
        var dir = Path.Combine(Path.GetTempPath(), "DesktopAutomationHelper", "Recordings");
        Directory.CreateDirectory(dir);

        var stamp = (_stoppedAt ?? DateTimeOffset.UtcNow).ToString("yyyyMMdd_HHmmss");
        _exportFilePath = Path.Combine(dir, $"recording_{stamp}.json");

        List<RecordedAction> snapshot;
        lock (_lock) { snapshot = [.. _actions]; }

        var export = new RecordingExport
        {
            StartedAt = _startedAt,
            StoppedAt = _stoppedAt,
            Mode = _currentMode.ToString(),
            ExportedFilePath = _exportFilePath,
            Actions = snapshot
        };

        File.WriteAllText(_exportFilePath, JsonSerializer.Serialize(export, JsonOpts));
    }

    private RecordingExport BuildExport()
    {
        List<RecordedAction> snapshot;
        lock (_lock) { snapshot = [.. _actions]; }

        return new RecordingExport
        {
            StartedAt = _startedAt,
            StoppedAt = _stoppedAt,
            Mode = _currentMode.ToString(),
            ExportedFilePath = _exportFilePath,
            Actions = snapshot
        };
    }
}
