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
    /// Optional low-level operation name to use during playback/export
    /// (for example <c>clicklogicalmenupath</c>).
    /// </summary>
    public string? Operation { get; set; }

    /// <summary>
    /// Optional logical menu path captured for menu-path actions.
    /// Entries are ordered from parent to target item.
    /// </summary>
    public List<ElementInfo>? MenuPath { get; set; }

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

    /// <summary>Additional operation-specific values captured for playback/export.</summary>
    public Dictionary<string, string>? Metadata { get; set; }

    /// <summary>Deterministic Assistive event id such as <c>evt-000001</c>.</summary>
    public string? EventId { get; set; }

    /// <summary>1-based Assistive event sequence within the recording session.</summary>
    public int? Sequence { get; set; }

    /// <summary>Canonical Jira key when the action was recorded inside a Jira scope.</summary>
    public string? JiraKey { get; set; }

    /// <summary>Optional BDD association. Omitted when no BDD is active for this action.</summary>
    public RecordedBddAssociation? Bdd { get; set; }

    /// <summary>Top-level window context captured before Assistive action execution.</summary>
    public RecordedWindowContext? Window { get; set; }

    /// <summary>Deterministic page id derived from the window title.</summary>
    public string? PageId { get; set; }

    /// <summary>Phase-3-style object reference such as <c>welcome.abc</c>.</summary>
    public string? ObjectRef { get; set; }

    /// <summary>Object reference for a drag-and-drop target element.</summary>
    public string? TargetObjectRef { get; set; }
}
