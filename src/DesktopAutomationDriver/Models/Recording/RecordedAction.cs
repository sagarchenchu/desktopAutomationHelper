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

    /// <summary>
    /// Human-readable description of the complete action, combining the action
    /// type and the element it was performed on (e.g. "Click on Login Button",
    /// "Is Visible check on Submit Button: True").
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// The text value associated with the action (e.g. the string typed into an Edit field).
    /// Null for actions that do not carry a value.
    /// </summary>
    public string? Value { get; set; }

    /// <summary>
    /// The drop-target element for a <see cref="ActionType.DragAndDrop"/> action.
    /// <see cref="Element"/> holds the drag source; this property holds the drop destination.
    /// Null for all other action types.
    /// </summary>
    public ElementInfo? TargetElement { get; set; }

    /// <summary>
    /// Mouse coordinate diagnostics captured while resolving the action target.
    /// Primarily populated for assistive and right-click-driven actions.
    /// </summary>
    public PointerContextInfo? PointerContext { get; set; }
}
