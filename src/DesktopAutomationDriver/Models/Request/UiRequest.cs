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

    /// <summary>
    /// When true, XPath is evaluated before attribute-based conditions.
    /// Preserves the legacy XPath-first search order even for fast operations.
    /// </summary>
    public bool? PreferXPath { get; set; }

    /// <summary>
    /// When true, only XPath is used for element lookup; attribute-based search is skipped.
    /// The operation will fail if the locator has no XPath expression.
    /// </summary>
    public bool? XPathOnly { get; set; }

    /// <summary>
    /// When true, attribute-based conditions (AutomationId, Name, ClassName, ControlType) are
    /// tried before XPath. This is the default for fast element operations.
    /// </summary>
    public bool? PreferAttributes { get; set; }

    /// <summary>
    /// Optional parent container locator. When set, the parent element is resolved first
    /// and the primary <see cref="Locator"/> is searched within that parent scope,
    /// reducing the search space and improving performance.
    /// </summary>
    public UiLocator? ParentLocator { get; set; }

    /// <summary>
    /// When true and <see cref="ParentLocator"/> is set, a failed child lookup inside
    /// the parent will be retried against the full window root before giving up.
    /// Defaults to <c>false</c> so that <see cref="ParentLocator"/> always narrows the
    /// search scope for speed.
    /// </summary>
    public bool? FallbackToWindowRootIfParentChildNotFound { get; set; }

    // -------------------------------------------------------------------------
    // switchwindow extended matching
    // -------------------------------------------------------------------------

    /// <summary>
    /// Native window handle for an exact HWND match in 'switchwindow'.
    /// Passed as a JSON number; stored as <c>long</c> to accommodate 64-bit handles.
    /// </summary>
    public long? Hwnd { get; set; }

    /// <summary>
    /// Win32 class name filter for 'switchwindow'.
    /// Matched case-insensitively; a substring match is also accepted.
    /// </summary>
    public string? ClassName { get; set; }

    /// <summary>
    /// Optional process ID override for 'switchwindow'.
    /// When provided it replaces the session application PID for the Win32 search.
    /// </summary>
    public int? ProcessId { get; set; }

    /// <summary>
    /// Title matching mode for 'switchwindow'.
    /// Accepted values: <c>exact</c>, <c>contains</c> (default), <c>regex</c>.
    /// </summary>
    public string? MatchMode { get; set; }

    /// <summary>
    /// When true, the 'quit' operation force-kills the application process tree even if
    /// this session was attached to an already-running process (not launched by the driver).
    /// Use with caution: this kills the process unconditionally.
    /// Defaults to false; only launched sessions are killed by default.
    /// </summary>
    public bool ForceKillAttachedProcess { get; set; }
}
