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
    /// Safe to call even if recording was already stopped (e.g. via Ctrl+S).
    /// Does not depend on the overlay existing or closing successfully.
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
    /// Lightweight point lookup for assistive status updates and popup probes.
    /// Must be called from an STA thread.
    /// </summary>
    ElementInfo? GetElementAtPointLightweight(System.Drawing.Point point);

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
    /// Used by Assistive-mode diagnostics, bounds fallback, and Ctrl+Right-Click
    /// handling to distinguish the application window from foreground popups.
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

    // ---- Assistive Jira / BDD annotation (session metadata; not RecordedActions) ----

    /// <summary>Stable recording id for the active session, if any.</summary>
    string? RecordingId { get; }

    /// <summary>Canonical Jira key for the active Assistive Jira scope, if set.</summary>
    string? JiraKey { get; }

    /// <summary>Safe overlay status fragment for Jira/BDD state (no statement text).</summary>
    string AssistiveAnnotationStatus { get; }

    /// <summary>Starts or replaces the Jira key before any Jira-scoped action is recorded.</summary>
    bool TryStartJiraRecording(string? rawKey, out string canonical, out string error);

    /// <summary>Arms a BDD statement for the next action or until finished.</summary>
    bool TryArmBddStatement(string? statement, BddScope scope, out string groupId, out string error);

    /// <summary>Finishes a multiple-action BDD group so later actions are not associated.</summary>
    bool TryFinishBddStatement(out string message);

    /// <summary>Cancels a pending BDD group that has not yet associated any actions.</summary>
    bool TryCancelBddStatement(out string message);

    /// <summary>
    /// Central Assistive recording path: assigns event/sequence, Jira/BDD, window/page/object refs.
    /// Every Assistive action must provide capture context from the target element/window.
    /// </summary>
    void RecordAssistiveAction(RecordedAction action, AssistiveActionCaptureContext captureContext);
}
