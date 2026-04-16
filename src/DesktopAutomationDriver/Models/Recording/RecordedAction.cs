namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Represents a single recorded automation action.
/// </summary>
public class RecordedAction
{
    /// <summary>The kind of action that was recorded.</summary>
    public ActionType ActionType { get; set; }

    /// <summary>UTC time the action was captured.</summary>
    public DateTimeOffset Timestamp { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>The recording mode that was active when the action was captured.</summary>
    public RecordingMode Mode { get; set; }

    /// <summary>Information about the UI element involved in the action.</summary>
    public ElementInfo? Element { get; set; }

    /// <summary>
    /// Result of query actions (IsVisible, IsClickable, IsEnabled, IsDisabled).
    /// Null for interactive actions.
    /// </summary>
    public bool? QueryResult { get; set; }
}
