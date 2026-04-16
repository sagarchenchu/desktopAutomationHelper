namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Describes the active recording mode.
/// </summary>
public enum RecordingMode
{
    /// <summary>No recording is active; waiting for the user to select a mode.</summary>
    None,

    /// <summary>
    /// Passive recording: mouse clicks and keyboard actions are captured
    /// automatically together with the element under the cursor.
    /// </summary>
    Passive,

    /// <summary>
    /// Assistive (armed) recording: the user right-clicks an element to choose
    /// which action to record/perform from a context menu.
    /// </summary>
    Assistive
}
