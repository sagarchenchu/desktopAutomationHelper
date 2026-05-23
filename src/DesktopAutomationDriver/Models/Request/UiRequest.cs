namespace DesktopAutomationDriver.Models.Request;

/// <summary>
/// Unified request body for the POST /ui endpoint.
/// </summary>
public class UiRequest
{
    /// <summary>
    /// The operation to perform (e.g. "launch", "click", "gettext").
    /// Case-insensitive.
    /// </summary>
    public string Operation { get; set; } = string.Empty;

    /// <summary>
    /// Primary element locator. Required by element-level operations.
    /// Multiple properties are combined with AND logic.
    /// </summary>
    public UiLocator? Locator { get; set; }

    /// <summary>
    /// Secondary element locator. Used by position-comparison operations
    /// (isrightof, isleftof, isabove, isbelow, getposition).
    /// </summary>
    public UiLocator? Locator2 { get; set; }

    /// <summary>
    /// Auxiliary string value whose meaning depends on the operation:
    /// launch → exe path; switchwindow → partial window title;
    /// type → text to type; sendkeys → key sequence; waitfor → timeout seconds;
    /// screenshot → optional file path; select → item name; etc.
    /// </summary>
    public string? Value { get; set; }

    /// <summary>
    /// Optional header dropdown click region for header dropdown operations.
    /// Defaults to LowerRight when omitted. Supports corner regions and right icon slots.
    /// </summary>
    public string? ClickRegion { get; set; }

    /// <summary>
    /// Optional dropdown ListItem click region for header dropdown item selection.
    /// Defaults to LeftCenter when omitted.
    /// </summary>
    public string? ItemRegion { get; set; }

    /// <summary>
    /// Zero-based index used by the "select" operation to pick a ComboBox item by position,
    /// and by "clickGridCell" to specify the row index.
    /// </summary>
    public int? Index { get; set; }

    /// <summary>
    /// Zero-based column index used by the "clickGridCell" operation.
    /// </summary>
    public int? ColumnIndex { get; set; }

    /// <summary>
    /// Optional maximum number of items returned by list-style operations.
    /// </summary>
    public int? Limit { get; set; }

    /// <summary>
    /// When true, listwindows also scans desktop descendants. Defaults to false.
    /// </summary>
    public bool? IncludeDesktopDescendants { get; set; }

    /// <summary>
    /// When true, ComboBox selection may use keyboard type-ahead fallback after
    /// bounded visible-list search fails. Defaults to false for huge ComboBoxes.
    /// </summary>
    public bool? AllowKeyboardFallback { get; set; }

    /// <summary>
    /// Overrides the default operation timeout in milliseconds.
    /// When set, the policy timeout is derived from this value regardless of the operation type.
    /// </summary>
    public int? TimeoutMs { get; set; }

    /// <summary>
    /// When true, forces fast-path behavior: short timeout (100 ms retry interval),
    /// no desktop popup scanning, and element caching enabled.
    /// </summary>
    public bool? Fast { get; set; }

    /// <summary>
    /// When true, disables the automatic popup/dialog window follow logic for this operation.
    /// Useful for operations where popup detection is not needed and scanning would add latency.
    /// </summary>
    public bool? DisableAutoFollow { get; set; }

    /// <summary>
    /// When true, enables element locator caching for this operation.
    /// A previously found element matching the locator will be returned from cache without re-scanning.
    /// </summary>
    public bool? UseCache { get; set; }
}
