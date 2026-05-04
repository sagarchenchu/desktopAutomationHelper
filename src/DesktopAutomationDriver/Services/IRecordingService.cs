using DesktopAutomationDriver.Models.Recording;
using DesktopAutomationDriver.Models.Request;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Controls the window-activity recording session.
/// </summary>
public interface IRecordingService
{
    /// <summary>True when a recording overlay is open and a mode is active or pending selection.</summary>
    bool IsActive { get; }

    /// <summary>The recording mode currently selected (None until the user presses Ctrl+P or Ctrl+A).</summary>
    RecordingMode CurrentMode { get; }

    /// <summary>UTC time the recording was started.</summary>
    DateTimeOffset? StartedAt { get; }

    /// <summary>
    /// Opens the transparent recording overlay and installs low-level input hooks.
    /// Optionally launches an application and schedules an automatic stop.
    /// Returns a <see cref="StartRecordingResult"/> whose <c>Error</c> is non-null on failure.
    /// </summary>
    StartRecordingResult StartRecording(StartRecordingRequest? request = null);

    /// <summary>
    /// Stops the active recording, writes the JSON export file and returns the result.
    /// Safe to call even if recording has already stopped.
    /// </summary>
    RecordingExport StopRecording();

    /// <summary>Returns the current state (including all recorded actions so far).</summary>
    RecordingExport GetCurrentState();

    // ---- called from the overlay window ----

    /// <summary>Changes the recording mode and updates the overlay display.</summary>
    void SetMode(RecordingMode mode);

    /// <summary>Appends a recorded action to the session log.</summary>
    void AddAction(RecordedAction action);

    /// <summary>
    /// Replaces the most recently recorded action with <paramref name="replacement"/>.
    /// Used by passive-mode drag detection to upgrade a prematurely recorded Click into a
    /// DragAndDrop action once a significant mouse movement has been observed.
    /// If the action log is empty the replacement is appended instead.
    /// </summary>
    void ReplaceLastAction(RecordedAction replacement);

    /// <summary>
    /// Uses UI Automation to identify the element at the given screen point.
    /// Must be called from an STA thread.
    /// </summary>
    ElementInfo? GetElementAtPoint(System.Drawing.Point point);

    /// <summary>
    /// Returns element information together with its immediate children names.
    /// Must be called from an STA thread.
    /// </summary>
    (ElementInfo? info, IReadOnlyList<ElementInfo> children) GetElementWithChildrenAtPoint(System.Drawing.Point point);

    /// <summary>
    /// Notifies the service that the overlay has closed (e.g. via Ctrl+S).
    /// Triggers JSON export if not already done.
    /// </summary>
    void OnOverlayClosed();

    /// <summary>
    /// Attempts to bring the active session's application window to the foreground
    /// so that elements in that window are at the top of the z-order and can be
    /// correctly identified via UIA <c>FromPoint</c>.
    /// Safe to call when no session is active (no-op).
    /// </summary>
    void BringApplicationWindowToFront();

    /// <summary>
    /// Returns the native window handle (HWND) of the active session's application
    /// main window, or <see cref="IntPtr.Zero"/> when no session is active or the
    /// handle cannot be obtained.
    /// Used by the Assistive-mode Ctrl+Right-Click handler to distinguish a
    /// foreground popup window from the application's own main window.
    /// </summary>
    IntPtr GetApplicationMainWindowHandle();

    /// <summary>
    /// Returns the process ID of the recording target application (the app launched or
    /// attached when <see cref="StartRecording"/> was called), or <c>null</c> when unknown.
    /// </summary>
    int? GetRecordingTargetProcessId();

    /// <summary>
    /// Returns the main window handle of the recording target application, refreshing it
    /// from the process when the stored handle is no longer valid.
    /// Returns <see cref="IntPtr.Zero"/> when the target is unknown or the window cannot
    /// be found.
    /// </summary>
    IntPtr GetRecordingTargetMainWindowHandle();

    /// <summary>
    /// Returns <c>true</c> when <paramref name="element"/> belongs to the recording target
    /// process/window. When no target is known, always returns <c>true</c> to preserve
    /// legacy behaviour.
    /// </summary>
    bool IsElementInRecordingTarget(AutomationElement element);

    /// <summary>
    /// Updates the runtime recording target to the given window handle.
    /// Optionally accepts a process ID; if omitted, the existing target PID is kept.
    /// Adds the HWND to the set of allowed target windows so that later
    /// <see cref="IsElementInRecordingTarget"/> checks accept elements inside it.
    /// </summary>
    void SetRecordingTargetWindow(IntPtr hwnd, int? processId = null, string? reason = null);
}
