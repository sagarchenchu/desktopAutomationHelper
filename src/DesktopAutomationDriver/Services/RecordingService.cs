using System.Diagnostics;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Request;
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

    // Custom output path supplied by the caller; null → use default temp directory
    private string? _outputPath;

    // Auto-stop timer (fires when waitSeconds elapses)
    private System.Threading.Timer? _autoStopTimer;

    private Thread? _overlayThread;
    private RecordingOverlayWindow? _overlayWindow;

    // Used by the overlay to run FromPoint on the STA thread
    private UIA3Automation? _automation;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
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

    public StartRecordingResult StartRecording(StartRecordingRequest? request = null)
    {
        if (_isActive)
            return new StartRecordingResult { Error = "Recording is already active." };

        lock (_lock)
        {
            _actions.Clear();
            _exportFilePath = null;
            _stoppedAt = null;
            _currentMode = RecordingMode.None;
            _outputPath = request?.OutputPath;
            _startedAt = DateTimeOffset.UtcNow;
            _isActive = true;
        }

        // ── Optional: launch the target application ──────────────────────────
        LaunchInfo? launchInfo = null;
        if (!string.IsNullOrWhiteSpace(request?.ExePath))
            launchInfo = LaunchApplication(request.ExePath);

        // ── Start the overlay on a dedicated STA thread ───────────────────────
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

            // Show the "stopped" notification in the top-right corner on this STA thread
            var notif = new RecordingStoppedNotification(_exportFilePath);
            WinForms.Application.Run(notif);
        })
        {
            IsBackground = true,
            Name = "RecordingOverlay-STA"
        };
        _overlayThread.SetApartmentState(ApartmentState.STA);
        _overlayThread.Start();

        // ── Optional: schedule auto-stop after waitSeconds ────────────────────
        if (request?.WaitSeconds is > 0)
        {
            const int MillisecondsPerSecond = 1000;
            var ms = request.WaitSeconds.Value * MillisecondsPerSecond;
            _autoStopTimer = new System.Threading.Timer(_ =>
            {
                _autoStopTimer?.Dispose();
                _autoStopTimer = null;
                StopRecording();
            }, null, ms, Timeout.Infinite);
        }

        _logger.LogInformation("Recording session started at {Time}", _startedAt);

        // Resolve the output path to include in the response
        var outputPath = ResolveOutputDirectory();

        return new StartRecordingResult
        {
            Launch = launchInfo,
            OutputPath = outputPath
        };
    }

    public RecordingExport StopRecording()
    {
        // Cancel any pending auto-stop timer
        _autoStopTimer?.Dispose();
        _autoStopTimer = null;

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

    public void ReplaceLastAction(RecordedAction replacement)
    {
        replacement.Timestamp = DateTimeOffset.UtcNow;
        replacement.Mode = _currentMode;

        if (string.IsNullOrEmpty(replacement.Description))
        {
            var elementLabel = ElementInfo.GetLabel(replacement.Element);
            replacement.Description = $"{replacement.ActionType} on {elementLabel}";
        }

        lock (_lock)
        {
            if (_actions.Count > 0)
                _actions[^1] = replacement;
            else
                _actions.Add(replacement);
        }

        _logger.LogDebug("Replaced last action with: {Type} [{Source}] → [{Target}]",
            replacement.ActionType,
            replacement.Element?.ControlType ?? "?",
            replacement.TargetElement?.ControlType ?? "?");
    }

    public ElementInfo? GetElementAtPoint(System.Drawing.Point point)
    {
        if (_automation == null) return null;
        try
        {
            var element = _automation.FromPoint(point);
            if (element == null) return null;
            element = RecordingOverlayWindow.DrillDownToElementAtPoint(element, point);
            return RecordingOverlayWindow.BuildElementInfo(element);
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
        _autoStopTimer?.Dispose();
        _autoStopTimer = null;
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
        var dir = ResolveOutputDirectory();
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

    /// <summary>
    /// Returns the directory that will hold the exported JSON file.
    /// If the caller supplied an OutputPath it is used directly (treated as a directory).
    /// Otherwise, falls back to %TEMP%\DesktopAutomationHelper\Recordings\.
    /// </summary>
    private string ResolveOutputDirectory()
    {
        if (!string.IsNullOrWhiteSpace(_outputPath))
        {
            // Treat the supplied path as a directory (create it if it doesn't exist)
            return _outputPath;
        }
        return Path.Combine(Path.GetTempPath(), "DesktopAutomationHelper", "Recordings");
    }

    /// <summary>
    /// Launches the application at <paramref name="exePath"/> and returns launch details.
    /// A brief wait allows the main window title to become available.
    /// </summary>
    private LaunchInfo LaunchApplication(string exePath)
    {
        try
        {
            var psi = new ProcessStartInfo(exePath) { UseShellExecute = true };
            var process = Process.Start(psi);
            if (process == null)
                return new LaunchInfo { Success = false, Error = "Process.Start returned null." };

            // Give the process up to 3 seconds to show a main window so we can read its title
            process.WaitForInputIdle(3000);

            string? title = null;
            for (int i = 0; i < 10; i++)
            {
                process.Refresh();
                title = process.MainWindowTitle;
                if (!string.IsNullOrEmpty(title)) break;
                Thread.Sleep(300);
            }

            // Sanitize user-provided values before logging to prevent log-forging
            var safeExe = SanitizeForLog(exePath);
            var safeTitle = SanitizeForLog(title ?? "(none)");
            _logger.LogInformation(
                "Launched application '{Exe}' as PID {Pid}, title '{Title}'",
                safeExe, process.Id, safeTitle);

            return new LaunchInfo
            {
                Success = true,
                ProcessId = process.Id,
                WindowTitle = title
            };
        }
        catch (Exception ex)
        {
            var safeExe = SanitizeForLog(exePath);
            _logger.LogError(ex, "Failed to launch application '{Exe}'", safeExe);
            return new LaunchInfo { Success = false, Error = ex.Message };
        }
    }

    /// <summary>
    /// Removes newline and carriage-return characters from a user-supplied string
    /// before it is written to a log entry, preventing log-injection attacks.
    /// </summary>
    private static string SanitizeForLog(string value) =>
        value.Replace("\r", string.Empty, StringComparison.Ordinal)
             .Replace("\n", string.Empty, StringComparison.Ordinal);

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
