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
    public bool ForceKillAttachedProcess { get; set; } = false;

    // -------------------------------------------------------------------------
    // Popup pipeline fields
    // -------------------------------------------------------------------------

    /// <summary>
    /// Popup action to perform.
    /// Accepted values: <c>button</c> (default), <c>close</c>, <c>enter</c>, <c>escape</c>, <c>makecurrent</c>.
    /// </summary>
    public string? Action { get; set; }

    /// <summary>
    /// Button name(s) to click for 'popupaction' with action=button.
    /// Pipe-separated list of candidates tried in order, e.g. <c>OK|Yes|Save</c>.
    /// </summary>
    public string? Button { get; set; }

    /// <summary>
    /// When true (default), popup discovery also scans the desktop root for windows
    /// outside the application process. Set to false to restrict search to app windows only.
    /// </summary>
    public bool? DesktopSearch { get; set; }

    /// <summary>
    /// When true, popup discovery only considers windows belonging to the current
    /// session application process. Defaults to false.
    /// </summary>
    public bool? SameProcessOnly { get; set; }

    /// <summary>
    /// When true (default for topwindow/waitforpopup/popupaction), the found popup
    /// is made the session's active window. Set to false to find without switching focus.
    /// </summary>
    public bool? MakeCurrent { get; set; }

    // -------------------------------------------------------------------------
    // dragbyoffset fields
    // -------------------------------------------------------------------------

    /// <summary>
    /// Horizontal movement in pixels for the 'dragbyoffset' operation.
    /// Defaults to 0.
    /// </summary>
    public int? OffsetX { get; set; }

    /// <summary>
    /// Vertical movement in pixels for the 'dragbyoffset' operation.
    /// Defaults to 0.
    /// </summary>
    public int? OffsetY { get; set; }

    /// <summary>
    /// Where to start the drag inside the element rectangle for 'dragbyoffset'.
    /// Supported values: center, topEdge, bottomEdge, leftEdge, rightEdge,
    /// topLeft, topRight, bottomLeft, bottomRight.
    /// Defaults to "center".
    /// </summary>
    public string? DragStart { get; set; }

    /// <summary>
    /// Total drag duration in milliseconds for the 'dragbyoffset' operation.
    /// Defaults to 250.
    /// </summary>
    public int? DragDurationMs { get; set; }

    /// <summary>
    /// Number of mouse-move steps for the 'dragbyoffset' operation.
    /// Defaults to 10.
    /// </summary>
    public int? DragSteps { get; set; }

    // -------------------------------------------------------------------------
    // dragcoordinates fields
    // -------------------------------------------------------------------------

    /// <summary>
    /// Source X screen coordinate for the 'dragcoordinates' operation.
    /// </summary>
    public int? FromX { get; set; }

    /// <summary>
    /// Source Y screen coordinate for the 'dragcoordinates' operation.
    /// </summary>
    public int? FromY { get; set; }

    /// <summary>
    /// Destination X screen coordinate for the 'dragcoordinates' operation.
    /// </summary>
    public int? ToX { get; set; }

    /// <summary>
    /// Destination Y screen coordinate for the 'dragcoordinates' operation.
    /// </summary>
    public int? ToY { get; set; }

    // -------------------------------------------------------------------------
    // mouse low-level fields
    // -------------------------------------------------------------------------

    /// <summary>
    /// Target X screen coordinate for the 'mouse' operation and coordinate-based scroll.
    /// </summary>
    public int? X { get; set; }

    /// <summary>
    /// Target Y screen coordinate for the 'mouse' operation and coordinate-based scroll.
    /// </summary>
    public int? Y { get; set; }

    // -------------------------------------------------------------------------
    // scroll fields
    // -------------------------------------------------------------------------

    /// <summary>
    /// Optional scrollable container locator for 'scroll' operations (Case C).
    /// When set together with <see cref="Locator"/>, the driver scrolls this container
    /// until the target element identified by <see cref="Locator"/> becomes visible.
    /// </summary>
    public UiLocator? ContainerLocator { get; set; }

    /// <summary>
    /// When true, element search for scroll-into-view operations includes off-screen elements.
    /// When null or true (the default), the implementation treats it as true so that
    /// virtualized items can be found before being scrolled into view.
    /// </summary>
    public bool? IncludeOffscreen { get; set; }

    /// <summary>
    /// Maximum number of scroll attempts before giving up in scroll-into-view loops.
    /// Defaults to 30.
    /// </summary>
    public int? MaxAttempts { get; set; }

    /// <summary>
    /// Delay in milliseconds between scroll attempts in scroll-into-view loops.
    /// Defaults to 150.
    /// </summary>
    public int? DelayMs { get; set; }

    /// <summary>
    /// Scroll direction for the 'scroll', 'mousescroll', 'wheelscroll' operations.
    /// Accepted values: up, down, left, right. Defaults to "down".
    /// </summary>
    public string? Direction { get; set; }

    /// <summary>
    /// Scroll amount (wheel ticks or pattern units) for scroll operations.
    /// Defaults to 1. Negative values reverse the direction.
    /// </summary>
    public int? Amount { get; set; }

    /// <summary>
    /// Scroll mode for the 'scroll' operation.
    /// Accepted values: auto (default), wheel, pattern.
    /// auto = try UIA ScrollPattern first, fall back to mouse wheel.
    /// wheel = physical mouse wheel only.
    /// pattern = UIA ScrollPattern only, fail if not supported.
    /// </summary>
    public string? Mode { get; set; }

    /// <summary>
    /// Raw wheel delta for 'mouse' action='scroll'. Negative = down/left, positive = up/right.
    /// When set, overrides Direction+Amount for the mouse scroll action.
    /// </summary>
    public int? WheelDelta { get; set; }

    /// <summary>
    /// When true, scroll attempts to verify that the scroll position actually changed.
    /// Defaults to false.
    /// </summary>
    public bool? VerifyScroll { get; set; }

    /// <summary>
    /// Milliseconds to wait after the scroll operation completes. Defaults to 100.
    /// </summary>
    public int? ScrollDelayMs { get; set; }

    // -------------------------------------------------------------------------
    // wait operation fields
    // -------------------------------------------------------------------------

    /// <summary>
    /// Target state for the 'wait' operation.
    /// Supported values: exists, enabled, disabled, visible, hidden, focused,
    /// windowactive, clickable, editable, gone.
    /// Defaults to "exists" when omitted.
    /// </summary>
    public string? State { get; set; }

    /// <summary>
    /// Polling interval in milliseconds for the 'wait' operation.
    /// Defaults to 200 ms.
    /// </summary>
    public int? PollIntervalMs { get; set; }

    // -------------------------------------------------------------------------
    // Central resolver search options (Phase 1.2)
    // -------------------------------------------------------------------------

    /// <summary>
    /// When true, the resolver returns all matching candidates instead of just the first.
    /// Used by the 'findall' operation and diagnostic queries.
    /// </summary>
    public bool? ReturnAllMatches { get; set; }

    /// <summary>
    /// Maximum number of candidates to return when <see cref="ReturnAllMatches"/> is true.
    /// Defaults to 500.
    /// </summary>
    public int? MaxMatches { get; set; }

    /// <summary>
    /// When true, the response includes a diagnostics block with candidate summaries,
    /// root strategy, errors, and candidate count even on success.
    /// </summary>
    public bool? IncludeDiagnostics { get; set; }

    /// <summary>
    /// When true, the resolver may return the best-scoring partial match when no exact match exists.
    /// Requires <see cref="BestMatch"/> to be set or the locator to supply enough hints for scoring.
    /// </summary>
    public bool? AllowBestMatch { get; set; }

    /// <summary>
    /// Name hint used for best-match scoring when <see cref="AllowBestMatch"/> is true.
    /// </summary>
    public string? BestMatch { get; set; }

    /// <summary>
    /// When true, the resolver uses the UIA desktop root as the search root,
    /// enabling cross-process and cross-window element lookup.
    /// </summary>
    public bool? UseDesktopRoot { get; set; }

    /// <summary>
    /// When true, the resolver uses the current active (foreground) window root
    /// regardless of the session's tracked active window.
    /// </summary>
    public bool? UseActiveWindowRoot { get; set; }

    /// <summary>
    /// When true, post-action verification failures do not throw; a warning is logged instead.
    /// Useful for dropdowns and custom controls where state transitions are unreliable.
    /// </summary>
    public bool? SoftVerification { get; set; }

}
