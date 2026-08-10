using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.Assistive;
using FlaUI.Core.AutomationElements;
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
    // Keep assistive point lookups responsive by returning only the first visible target children;
    // larger child sets are intentionally truncated for menu/status previews.
    private const int MaxChildrenAtPoint = 30;

    private readonly ILogger<RecordingService> _logger;
    private readonly IUiSessionContext _sessionContext;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindow(IntPtr hWnd);

    private const uint GA_ROOT = 2;
    private const int SW_SHOWNORMAL = 1;
    private const int SW_SHOWMINIMIZED = 2;
    private const int SW_SHOWMAXIMIZED = 3;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WINDOWPLACEMENT
    {
        public int Length;
        public int Flags;
        public int ShowCmd;
        public POINT MinPosition;
        public POINT MaxPosition;
        public RECT NormalPosition;
    }

    private readonly List<RecordedAction> _actions = [];
    private readonly object _lock = new();
    private readonly AssistiveRecordingCoordinator _assistive = new();
    private readonly AssistiveArtifactBuilder _artifactBuilder = new();
    private AssistiveArtifactWriter _artifactWriter = new();
    private RecordingArtifactsSummary? _artifactsSummary;
    private RecordingExportState _exportState = RecordingExportState.NotStarted;

    /// <summary>Test seam: replaces the default artifact writer (path-safety / failure injection).</summary>
    internal AssistiveArtifactWriter ArtifactWriterForTests
    {
        get => _artifactWriter;
        set => _artifactWriter = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Test seam: inspect export-completed flag.</summary>
    internal bool ExportCompletedForTests
    {
        get { lock (_lock) return _exportState == RecordingExportState.Completed; }
    }

    /// <summary>Test seam: inspect export state machine.</summary>
    internal RecordingExportState ExportStateForTests
    {
        get { lock (_lock) return _exportState; }
    }

    /// <summary>Test seam: inspect primary export path.</summary>
    internal string? ExportFilePathForTests => _exportFilePath;

    /// <summary>Test seam: invoked immediately before the primary recording atomic write.</summary>
    internal Action<string>? BeforePrimaryWriteForTests { get; set; }

    /// <summary>Test seam: invoked after sidecars succeed, before primary summary rewrite.</summary>
    internal Action? BeforePrimarySummaryRewriteForTests { get; set; }

    // ── Recording target (process/window to which Assistive mode is scoped) ──
    private int? _recordingTargetProcessId;
    private IntPtr _recordingTargetMainHwnd = IntPtr.Zero;
    private string? _recordingTargetExePath;
    private readonly HashSet<IntPtr> _allowedTargetWindows = new();

    private volatile bool _isActive;
    private volatile RecordingMode _currentMode = RecordingMode.None;

    private DateTimeOffset _startedAt;
    private DateTimeOffset? _stoppedAt;
    private string? _exportFilePath;

    // Custom output path supplied by the caller; null → use default temp directory
    private string? _outputPath;
    private ScreenResolutionInfo? _screenInfo;
    private LaunchInfo? _launchInfo;

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

    public RecordingService(ILogger<RecordingService> logger, IUiSessionContext sessionContext)
    {
        _logger = logger;
        _sessionContext = sessionContext;
    }

    // ── IRecordingService ────────────────────────────────────────────────────

    public bool IsActive => _isActive;

    public RecordingMode CurrentMode => _currentMode;

    public DateTimeOffset? StartedAt => _isActive ? _startedAt : null;

    public string? RecordingId
    {
        get { lock (_lock) return _assistive.RecordingId; }
    }

    public string? JiraKey
    {
        get { lock (_lock) return _assistive.JiraKey; }
    }

    public string AssistiveAnnotationStatus
    {
        get { lock (_lock) return _assistive.OverlayStatusSuffix; }
    }

    public StartRecordingResult StartRecording(StartRecordingRequest? request = null)
    {
        if (_isActive)
            return new StartRecordingResult { Error = "Recording is already active." };

        lock (_lock)
        {
            if (_exportState == RecordingExportState.InProgress)
            {
                return new StartRecordingResult
                {
                    Error = "Cannot start a new recording while the previous export is still in progress."
                };
            }

            _actions.Clear();
            _exportFilePath = null;
            _artifactsSummary = null;
            _exportState = RecordingExportState.NotStarted;
            _stoppedAt = null;
            _currentMode = RecordingMode.None;
            _outputPath = request?.OutputPath;
            _startedAt = DateTimeOffset.UtcNow;
            _isActive = true;
            _screenInfo = CaptureScreenResolution();
            _launchInfo = null;

            var recordingId = $"rec-{_startedAt:yyyyMMdd-HHmmss}-{Guid.NewGuid().ToString("N")[..8]}";
            _assistive.Reset(recordingId);

            // Reset recording target
            _recordingTargetProcessId = null;
            _recordingTargetMainHwnd = IntPtr.Zero;
            _recordingTargetExePath = null;
            _allowedTargetWindows.Clear();
        }

        // ── Capture target from active UiSession (if any) ─────────────────────
        var activeSession = _sessionContext.ActiveSession;
        if (activeSession != null)
        {
            _recordingTargetProcessId = activeSession.Application.ProcessId;
            try
            {
                var proc = Process.GetProcessById(activeSession.Application.ProcessId);
                proc.Refresh();
                _recordingTargetMainHwnd = proc.MainWindowHandle;
                if (_recordingTargetMainHwnd != IntPtr.Zero)
                    _allowedTargetWindows.Add(_recordingTargetMainHwnd);
            }
            catch
            {
                _recordingTargetMainHwnd = IntPtr.Zero;
            }
        }

        // ── Optional: launch the target application ──────────────────────────
        LaunchInfo? launchInfo = null;
        if (!string.IsNullOrWhiteSpace(request?.ExePath))
        {
            launchInfo = LaunchApplication(request.ExePath);
            _launchInfo = launchInfo;

            if (launchInfo?.Success == true && launchInfo.ProcessId > 0)
            {
                _recordingTargetProcessId = launchInfo.ProcessId;
                _recordingTargetExePath = request.ExePath;

                try
                {
                    var proc = Process.GetProcessById(launchInfo.ProcessId.Value);
                    proc.Refresh();
                    _recordingTargetMainHwnd = proc.MainWindowHandle;
                    if (_recordingTargetMainHwnd != IntPtr.Zero)
                        _allowedTargetWindows.Add(_recordingTargetMainHwnd);
                }
                catch
                {
                    _recordingTargetMainHwnd = IntPtr.Zero;
                }
            }
        }

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
            Screen = _screenInfo,
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
        if (_currentMode == RecordingMode.Assistive)
        {
            throw new InvalidOperationException(
                "Assistive actions must be recorded via RecordAssistiveAction with capture context.");
        }

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

    public void RecordAssistiveAction(RecordedAction action, AssistiveActionCaptureContext captureContext)
    {
        ArgumentNullException.ThrowIfNull(action);
        ArgumentNullException.ThrowIfNull(captureContext);

        action.Timestamp = DateTimeOffset.UtcNow;

        if (string.IsNullOrEmpty(action.Description))
        {
            var elementLabel = ElementInfo.GetLabel(action.Element);
            action.Description = action.QueryResult.HasValue
                ? $"{action.ActionType} check on {elementLabel}: {action.QueryResult}"
                : $"{action.ActionType} on {elementLabel}";
        }

        lock (_lock)
        {
            _assistive.EnrichAssistiveAction(action, captureContext);
            _actions.Add(action);

            if (action.Bdd != null)
            {
                _logger.LogInformation(
                    "Assistive event {EventId} associated with BDD {GroupId} ({CharCount} chars).",
                    action.EventId,
                    action.Bdd.GroupId,
                    action.Bdd.Statement.Length);
            }
            else
            {
                _logger.LogDebug(
                    "Assistive event {EventId} recorded without BDD. pageId={PageId}",
                    action.EventId,
                    action.PageId ?? "(none)");
            }
        }
    }

    /// <summary>
    /// Test seam: configures an Assistive session without starting the overlay.
    /// </summary>
    internal void ConfigureAssistiveSessionForTests(string outputDirectory, string recordingId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        ArgumentException.ThrowIfNullOrWhiteSpace(recordingId);

        lock (_lock)
        {
            _actions.Clear();
            _exportFilePath = null;
            _artifactsSummary = null;
            _exportState = RecordingExportState.NotStarted;
            _outputPath = outputDirectory;
            _startedAt = DateTimeOffset.UtcNow;
            _stoppedAt = null;
            _isActive = true;
            _currentMode = RecordingMode.Assistive;
            _assistive.Reset(recordingId);
        }
    }

    /// <summary>Test seam: runs the export pipeline (idempotent / retry-safe).</summary>
    internal void ExportForTests()
    {
        lock (_lock)
        {
            _stoppedAt ??= DateTimeOffset.UtcNow;
            _isActive = false;
        }

        ExportJson();
    }

    public bool TryStartJiraRecording(string? rawKey, out string canonical, out string error)
    {
        lock (_lock)
        {
            var ok = _assistive.TryStartJiraRecording(rawKey, out canonical, out error);
            if (ok)
            {
                _logger.LogInformation(
                    "Jira recording scope set to {JiraKey}. Existing ordinary actions are preserved.",
                    canonical);
            }

            return ok;
        }
    }

    public bool TryArmBddStatement(string? statement, BddScope scope, out string groupId, out string error)
    {
        lock (_lock)
        {
            var ok = _assistive.TryArmBdd(statement, scope, out groupId, out error);
            if (ok)
            {
                _logger.LogInformation(
                    "BDD {GroupId} armed ({Scope}, {CharCount} chars).",
                    groupId,
                    scope,
                    statement?.Trim().Length ?? 0);
            }

            return ok;
        }
    }

    public bool TryFinishBddStatement(out string message)
    {
        lock (_lock)
        {
            var ok = _assistive.TryFinishBdd(out message);
            if (ok)
                _logger.LogInformation("{Message}", message);
            return ok;
        }
    }

    public bool TryCancelBddStatement(out string message)
    {
        lock (_lock)
        {
            var ok = _assistive.TryCancelBdd(out message);
            if (ok)
                _logger.LogInformation("{Message}", message);
            return ok;
        }
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
        using (MeasurePerf("GetElementAtPoint"))
        {
            return GetElementAtPointCore(point, logOutsideTarget: true);
        }
    }

    public ElementInfo? GetElementAtPointLightweight(System.Drawing.Point point)
    {
        using (MeasurePerf("GetElementAtPointLightweight"))
        {
            if (_automation == null)
                return null;

            try
            {
                var element = _automation.FromPoint(point);

                if (element == null)
                    return null;

                element = RecordingOverlayWindow.DrillDownToElementAtPoint(element, point);

                if (!IsElementInRecordingTarget(element))
                    return null;

                return RecordingOverlayWindow.BuildElementInfo(element);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "GetElementAtPointLightweight failed for {Point}", point);
                return null;
            }
        }
    }

    private ElementInfo? GetElementAtPointCore(System.Drawing.Point point, bool logOutsideTarget)
    {
        if (_automation == null) return null;
        try
        {
            var element = _automation.FromPoint(point);
            if (element == null) return null;
            element = RecordingOverlayWindow.DrillDownToElementAtPoint(element, point);

            if (!IsElementInRecordingTarget(element))
            {
                if (logOutsideTarget)
                {
                    _logger.LogWarning(
                        "Ignoring element outside recording target at {Point}: name={Name}, automationId={AutomationId}, processId={ProcessId}",
                        point,
                        SafeName(element),
                        SafeAutomationId(element),
                        SafeProcessId(element));
                }
                return null;
            }

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
            element = RecordingOverlayWindow.DrillDownToElementAtPoint(element, point);

            if (!IsElementInRecordingTarget(element))
                return (null, Array.Empty<ElementInfo>());

            var info = RecordingOverlayWindow.BuildElementInfo(element);
            var childInfos = element.FindAllChildren()
                .Where(IsElementInRecordingTarget)
                .Take(MaxChildrenAtPoint)
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

        // Export JSON (idempotent)
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

    private IDisposable MeasurePerf(string name)
    {
        return new PerfScope(_logger, name);
    }

    private sealed class PerfScope : IDisposable
    {
        private readonly ILogger _logger;
        private readonly string _name;
        private readonly Stopwatch _sw;

        public PerfScope(ILogger logger, string name)
        {
            _logger = logger;
            _name = name;
            _sw = Stopwatch.StartNew();
        }

        public void Dispose()
        {
            _sw.Stop();

            if (_sw.ElapsedMilliseconds > 250)
            {
                _logger.LogWarning(
                    "PERF: {Name} took {ElapsedMs} ms",
                    _name,
                    _sw.ElapsedMilliseconds);
            }
        }
    }

    public void Dispose()
    {
        _autoStopTimer?.Dispose();
        _autoStopTimer = null;
        CloseOverlayIfOpen();
        _automation?.Dispose();
    }

    /// <inheritdoc/>
    public void BringApplicationWindowToFront()
    {
        try
        {
            var hwnd = GetRecordingTargetMainWindowHandle();

            if (hwnd != IntPtr.Zero)
            {
                SetForegroundWindow(hwnd);
                return;
            }

            var session = _sessionContext.ActiveSession;
            if (session == null) return;

            var pid = session.Application.ProcessId;
            var proc = Process.GetProcessById(pid);
            hwnd = proc.MainWindowHandle;
            if (hwnd != IntPtr.Zero)
                SetForegroundWindow(hwnd);
        }
        catch { /* best effort — never fail recording because of a focus hint */ }
    }

    /// <inheritdoc/>
    public IntPtr GetApplicationMainWindowHandle()
    {
        var recordingTarget = GetRecordingTargetMainWindowHandle();
        if (recordingTarget != IntPtr.Zero)
            return recordingTarget;

        try
        {
            var session = _sessionContext.ActiveSession;
            if (session == null) return IntPtr.Zero;

            var pid = session.Application.ProcessId;
            var proc = Process.GetProcessById(pid);
            return proc.MainWindowHandle;
        }
        catch
        {
            return IntPtr.Zero;
        }
    }

    /// <inheritdoc/>
    public int? GetRecordingTargetProcessId()
    {
        if (_recordingTargetProcessId.HasValue)
            return _recordingTargetProcessId;

        try
        {
            var session = _sessionContext.ActiveSession;
            return session?.Application.ProcessId;
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public IntPtr GetRecordingTargetMainWindowHandle()
    {
        if (_recordingTargetMainHwnd != IntPtr.Zero && IsWindow(_recordingTargetMainHwnd))
            return _recordingTargetMainHwnd;

        try
        {
            var pid = GetRecordingTargetProcessId();
            if (pid.HasValue)
            {
                var proc = Process.GetProcessById(pid.Value);
                proc.Refresh();
                if (proc.MainWindowHandle != IntPtr.Zero)
                {
                    _recordingTargetMainHwnd = proc.MainWindowHandle;
                    return _recordingTargetMainHwnd;
                }
            }
        }
        catch
        {
            // best effort
        }

        return IntPtr.Zero;
    }

    /// <inheritdoc/>
    public bool IsElementInRecordingTarget(AutomationElement element)
    {
        if (element == null)
            return false;

        var targetPid = GetRecordingTargetProcessId();

        // Primary check: match by process ID
        try
        {
            var elementPid = element.Properties.ProcessId.Value;
            if (targetPid.HasValue && elementPid == targetPid.Value)
                return true;
        }
        catch
        {
            // continue to HWND fallback
        }

        // HWND fallback: check against allowed target windows and root HWND
        try
        {
            var hwndRaw = element.Properties.NativeWindowHandle.ValueOrDefault;
            if (hwndRaw != 0)
            {
                var hwnd = new IntPtr(hwndRaw);
                var root = GetAncestor(hwnd, GA_ROOT);

                if (_allowedTargetWindows.Contains(hwnd) ||
                    (root != IntPtr.Zero && _allowedTargetWindows.Contains(root)))
                {
                    return true;
                }

                var targetRoot = GetRecordingTargetMainWindowHandle();
                if (root != IntPtr.Zero && targetRoot != IntPtr.Zero && root == targetRoot)
                    return true;
            }
        }
        catch
        {
            // best effort
        }

        // No target known — preserve legacy behaviour
        if (!targetPid.HasValue && _allowedTargetWindows.Count == 0)
            return true;

        return false;
    }

    /// <inheritdoc/>
    public void SetRecordingTargetWindow(IntPtr hwnd, int? processId = null, string? reason = null)
    {
        if (hwnd == IntPtr.Zero)
            return;

        try
        {
            _recordingTargetMainHwnd = hwnd;
            _allowedTargetWindows.Add(hwnd);

            if (processId.HasValue)
                _recordingTargetProcessId = processId;

            _logger.LogInformation(
                "Recording target window updated. hwnd=0x{Hwnd:X}, pid={Pid}, reason={Reason}",
                hwnd.ToInt64(),
                _recordingTargetProcessId,
                reason ?? "(none)");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to update recording target window to 0x{Hwnd:X}", hwnd.ToInt64());
        }
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
        lock (_lock)
        {
            while (_exportState == RecordingExportState.InProgress)
                Monitor.Wait(_lock);

            if (_exportState == RecordingExportState.Completed)
                return;

            _exportState = RecordingExportState.InProgress;
        }

        try
        {
            ExportJsonCore();

            lock (_lock)
            {
                _exportState = RecordingExportState.Completed;
                Monitor.PulseAll(_lock);
            }
        }
        catch
        {
            // Primary write failed before a recoverable artifact outcome — allow retry.
            lock (_lock)
            {
                if (_exportState == RecordingExportState.InProgress)
                    _exportState = RecordingExportState.NotStarted;
                Monitor.PulseAll(_lock);
            }

            throw;
        }
    }

    private void ExportJsonCore()
    {
        var dir = ResolveOutputDirectory();
        Directory.CreateDirectory(dir);
        AssistivePathSafety.EnsureWritablePathInside(dir, dir);

        var stamp = (_stoppedAt ?? DateTimeOffset.UtcNow).ToString("yyyyMMdd_HHmmss");
        string exportFilePath;
        lock (_lock)
        {
            // Retry must rewrite the same primary path so a failed export can recover.
            _exportFilePath ??= Path.Combine(dir, $"recording_{stamp}.json");
            exportFilePath = _exportFilePath;
        }

        AssistivePathSafety.EnsureWritablePathInside(exportFilePath, dir);

        List<RecordedAction> snapshot;
        string? recordingId;
        string? jiraKey;
        lock (_lock)
        {
            snapshot = [.. _actions];
            recordingId = _assistive.RecordingId;
            jiraKey = _assistive.JiraKey;
        }

        recordingId ??= $"rec-{stamp}-{Guid.NewGuid().ToString("N")[..8]}";
        var recordingFileName = Path.GetFileName(exportFilePath);

        var export = new RecordingExport
        {
            StartedAt = _startedAt,
            StoppedAt = _stoppedAt,
            Mode = _currentMode.ToString(),
            Screen = _screenInfo,
            Launch = _launchInfo,
            ExportedFilePath = exportFilePath,
            RecordingId = recordingId,
            Actions = snapshot
        };

        // Primary recording must succeed even if sidecars fail.
        BeforePrimaryWriteForTests?.Invoke(exportFilePath);
        AssistiveAtomicIO.ReplaceFileAtomic(
            exportFilePath,
            JsonSerializer.Serialize(export, JsonOpts));

        RecordingArtifactsSummary? artifactSummary = null;
        try
        {
            var build = _artifactBuilder.Build(
                recordingId,
                recordingFileName,
                _stoppedAt ?? DateTimeOffset.UtcNow,
                jiraKey,
                snapshot);

            if (build.Pages.Count > 0 || build.BddActionMap is not null)
            {
                artifactSummary = _artifactWriter.Write(
                    dir,
                    recordingFileName,
                    recordingId,
                    jiraKey,
                    build);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Assistive artifact export failed; primary recording was preserved.");
            artifactSummary = new RecordingArtifactsSummary
            {
                Warnings = ["Assistive artifact export failed. Primary recording was preserved."]
            };
        }

        if (artifactSummary is null)
            return;

        _artifactsSummary = artifactSummary;
        export.Artifacts = artifactSummary;

        try
        {
            BeforePrimarySummaryRewriteForTests?.Invoke();
            AssistiveAtomicIO.ReplaceFileAtomic(
                exportFilePath,
                JsonSerializer.Serialize(export, JsonOpts));
        }
        catch (Exception rewriteEx)
        {
            // Sidecars (when present) already exist; do not reclassify as artifact export failure.
            _logger.LogWarning(
                rewriteEx,
                "Assistive sidecars were written but rewriting the primary recording with the artifact summary failed. Primary contents from the initial write remain intact.");

            if (!artifactSummary.Warnings.Any(w =>
                    w.Contains("rewriting the primary", StringComparison.OrdinalIgnoreCase)))
            {
                artifactSummary.Warnings.Add(
                    "Assistive sidecars were written but the primary recording could not be updated with the artifact summary.");
            }

            _artifactsSummary = artifactSummary;
        }
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
            try
            {
                process.WaitForInputIdle(3000);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "WaitForInputIdle failed for '{Exe}'", SanitizeForLog(exePath));
            }

            string? title = null;
            IntPtr mainWindowHandle = IntPtr.Zero;
            for (int i = 0; i < 10; i++)
            {
                process.Refresh();
                title = process.MainWindowTitle;
                mainWindowHandle = process.MainWindowHandle;
                if (!string.IsNullOrEmpty(title) || mainWindowHandle != IntPtr.Zero) break;
                Thread.Sleep(300);
            }

            var windowInfo = CaptureApplicationWindow(mainWindowHandle, title);

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
                WindowTitle = title,
                Window = windowInfo
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

    private static string? SafeName(AutomationElement element)
    {
        try { return element.Name; }
        catch { return null; }
    }

    private static string? SafeAutomationId(AutomationElement element)
    {
        try { return element.AutomationId; }
        catch { return null; }
    }

    private static int? SafeProcessId(AutomationElement element)
    {
        try { return element.Properties.ProcessId.Value; }
        catch { return null; }
    }

    private RecordingExport BuildExport()
    {
        bool shouldExport;
        lock (_lock)
            shouldExport = _exportState != RecordingExportState.Completed && _stoppedAt is not null;

        if (shouldExport)
        {
            try
            {
                ExportJson();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to export recording JSON on retry");
            }
        }

        List<RecordedAction> snapshot;
        string? recordingId;
        lock (_lock)
        {
            snapshot = [.. _actions];
            recordingId = _assistive.RecordingId;
        }

        return new RecordingExport
        {
            StartedAt = _startedAt,
            StoppedAt = _stoppedAt,
            Mode = _currentMode.ToString(),
            Screen = _screenInfo,
            Launch = _launchInfo,
            ExportedFilePath = _exportFilePath,
            RecordingId = recordingId,
            Artifacts = _artifactsSummary,
            Actions = snapshot
        };
    }

    private ScreenResolutionInfo? CaptureScreenResolution()
    {
        try
        {
            var screen = WinForms.Screen.PrimaryScreen ?? WinForms.Screen.AllScreens.FirstOrDefault();
            if (screen == null)
                return null;

            return new ScreenResolutionInfo
            {
                DeviceName = screen.DeviceName,
                IsPrimary = screen.Primary,
                X = screen.Bounds.X,
                Y = screen.Bounds.Y,
                Width = screen.Bounds.Width,
                Height = screen.Bounds.Height,
                WorkingAreaX = screen.WorkingArea.X,
                WorkingAreaY = screen.WorkingArea.Y,
                WorkingAreaWidth = screen.WorkingArea.Width,
                WorkingAreaHeight = screen.WorkingArea.Height
            };
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to capture screen resolution before recording start");
            return null;
        }
    }

    private ApplicationWindowInfo? CaptureApplicationWindow(IntPtr hwnd, string? title)
    {
        if (hwnd == IntPtr.Zero)
            return null;

        try
        {
            if (!GetWindowRect(hwnd, out var rect))
                return null;

            var placement = new WINDOWPLACEMENT { Length = Marshal.SizeOf<WINDOWPLACEMENT>() };
            var hasPlacement = GetWindowPlacement(hwnd, ref placement);

            var windowRect = new System.Drawing.Rectangle(
                rect.Left,
                rect.Top,
                Math.Max(0, rect.Right - rect.Left),
                Math.Max(0, rect.Bottom - rect.Top));

            var screen = WinForms.Screen.FromHandle(hwnd);
            var screenBounds = screen.Bounds;
            var isFullScreen =
                windowRect.Left <= screenBounds.Left &&
                windowRect.Top <= screenBounds.Top &&
                windowRect.Right >= screenBounds.Right &&
                windowRect.Bottom >= screenBounds.Bottom;

            var showCmd = hasPlacement ? placement.ShowCmd : 0;

            return new ApplicationWindowInfo
            {
                Title = title,
                X = windowRect.X,
                Y = windowRect.Y,
                Width = windowRect.Width,
                Height = windowRect.Height,
                WindowState = showCmd switch
                {
                    SW_SHOWMAXIMIZED => "maximized",
                    SW_SHOWMINIMIZED => "minimized",
                    SW_SHOWNORMAL => "normal",
                    _ => "unknown"
                },
                IsMaximized = showCmd == SW_SHOWMAXIMIZED,
                IsMinimized = showCmd == SW_SHOWMINIMIZED,
                IsFullScreen = isFullScreen
            };
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to capture launched window details for '{Title}'", title ?? "(unknown)");
            return null;
        }
    }
}
