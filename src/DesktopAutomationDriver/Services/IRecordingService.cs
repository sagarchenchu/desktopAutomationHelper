using DesktopAutomationDriver.Models.Recording;

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
    /// Returns an error message if recording is already active.
    /// </summary>
    string StartRecording();

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
}
