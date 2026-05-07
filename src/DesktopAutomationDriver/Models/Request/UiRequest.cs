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
}
