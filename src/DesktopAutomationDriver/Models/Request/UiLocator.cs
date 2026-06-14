namespace DesktopAutomationDriver.Models.Request;

/// <summary>
/// Identifies a UI element for a /ui endpoint operation.
/// Multiple properties are combined with AND logic.
/// Supports pywinauto-style extended properties for fine-grained element resolution.
/// </summary>
public class UiLocator
{
    /// <summary>
    /// Optional locator mode hint for operations that support multiple strategies
    /// (for example visual vs logical menu path resolution).
    /// </summary>
    public string? Mode { get; set; }

    /// <summary>UIA Name property (element label).</summary>
    public string? Name { get; set; }

    /// <summary>Regex pattern to match on Name.</summary>
    public string? NameRegex { get; set; }

    /// <summary>UIA AutomationId property.</summary>
    public string? AutomationId { get; set; }

    /// <summary>Regex pattern to match on AutomationId.</summary>
    public string? AutomationIdRegex { get; set; }

    /// <summary>UIA ClassName property.</summary>
    public string? ClassName { get; set; }

    /// <summary>Regex pattern to match on ClassName.</summary>
    public string? ClassNameRegex { get; set; }

    /// <summary>
    /// UIA control type string, e.g. "Button", "Edit", "ComboBox", "CheckBox", "Text".
    /// </summary>
    public string? ControlType { get; set; }

    /// <summary>
    /// XPath-style expression used to locate a UI element by traversing the UIA tree.
    /// Supported syntax:
    /// <list type="bullet">
    ///   <item><description><c>//ControlType[@Attr='value']</c> – find first descendant of the given type with the attribute.</description></item>
    ///   <item><description><c>//*[@AutomationId='id']</c> – find first descendant of any type with the given AutomationId.</description></item>
    ///   <item><description><c>//Parent/Child[@Name='ok']</c> – two-step: find Parent descendant, then its direct-child Child.</description></item>
    ///   <item><description><c>//ControlType[@Name='a' and @AutomationId='b']</c> – multiple attribute predicates joined by <c>and</c>.</description></item>
    ///   <item><description><c>//ControlType[@Attr='value'][2]</c> – 1-based index to pick the n-th matching element.</description></item>
    /// </list>
    /// Supported attributes: <c>@Name</c>, <c>@AutomationId</c>, <c>@ClassName</c>, <c>@ControlType</c>.
    /// When <c>XPath</c> is set, all other locator properties are ignored.
    /// </summary>
    public string? XPath { get; set; }

    // -------------------------------------------------------------------------
    // pywinauto-style extended identification fields
    // -------------------------------------------------------------------------

    /// <summary>Native window handle (HWND). When set, the element is resolved directly from this handle.</summary>
    public long? Hwnd { get; set; }

    /// <summary>Win32 process ID owning the element.</summary>
    public int? ProcessId { get; set; }

    /// <summary>Win32 control ID (GetDlgCtrlID).</summary>
    public int? ControlId { get; set; }

    /// <summary>UIA FrameworkId string, e.g. "WPF", "Win32", "WinForm".</summary>
    public string? FrameworkId { get; set; }

    /// <summary>UIA RuntimeId as a string (e.g. "42.100448"). Used for exact-element pinning.</summary>
    public string? RuntimeId { get; set; }

    /// <summary>
    /// Match on ValuePattern.Value. Useful for Edit/ComboBox/Custom controls.
    /// Matched using <see cref="ValueMatchMode"/> (default: exact).
    /// </summary>
    public string? Value { get; set; }

    /// <summary>Regex pattern to match on ValuePattern.Value.</summary>
    public string? ValueRegex { get; set; }

    /// <summary>
    /// Match on TextPattern content or element Name when TextPattern is not available.
    /// Matched using <see cref="TextMatchMode"/> (default: exact).
    /// </summary>
    public string? Text { get; set; }

    // -------------------------------------------------------------------------
    // State filters
    // -------------------------------------------------------------------------

    /// <summary>When set, only elements with matching visibility are returned (true = visible, false = not visible).</summary>
    public bool? Visible { get; set; }

    /// <summary>When set, only elements with matching enabled state are returned.</summary>
    public bool? Enabled { get; set; }

    /// <summary>When set, only elements with matching offscreen state are returned (true = offscreen, false = on-screen).</summary>
    public bool? Offscreen { get; set; }

    // -------------------------------------------------------------------------
    // Match modes — override global MatchMode per property
    // -------------------------------------------------------------------------

    /// <summary>
    /// Default match mode applied to all string properties when no per-property mode is set.
    /// Accepted values: <c>exact</c> (default), <c>contains</c>, <c>startswith</c>, <c>regex</c>.
    /// </summary>
    public string? MatchMode { get; set; }

    /// <summary>Match mode override for the <see cref="Name"/> property.</summary>
    public string? NameMatchMode { get; set; }

    /// <summary>Match mode override for the <see cref="AutomationId"/> property.</summary>
    public string? AutomationIdMatchMode { get; set; }

    /// <summary>Match mode override for the <see cref="ClassName"/> property.</summary>
    public string? ClassNameMatchMode { get; set; }

    /// <summary>Match mode override for the <see cref="Value"/> property.</summary>
    public string? ValueMatchMode { get; set; }

    /// <summary>Match mode override for the <see cref="Text"/> property.</summary>
    public string? TextMatchMode { get; set; }

    // -------------------------------------------------------------------------
    // Indexing / disambiguation (pywinauto ctrl_index / found_index)
    // -------------------------------------------------------------------------

    /// <summary>
    /// Zero-based index applied after filtering and scoring. Equivalent to pywinauto <c>found_index</c>.
    /// When set, returns the n-th element from the scored+filtered candidate list.
    /// </summary>
    public int? FoundIndex { get; set; }

    /// <summary>
    /// Zero-based raw index applied before text filtering. Equivalent to pywinauto <c>ctrl_index</c>.
    /// When set, returns the n-th collected candidate without further filtering.
    /// </summary>
    public int? CtrlIndex { get; set; }

    // -------------------------------------------------------------------------
    // Search scope
    // -------------------------------------------------------------------------

    /// <summary>Maximum UIA tree depth to search. Defaults to 20 when not set.</summary>
    public int? Depth { get; set; }

    /// <summary>When true, only direct children of the root are considered (no descendant traversal).</summary>
    public bool? TopLevelOnly { get; set; }

    /// <summary>When true, only elements in the active (foreground) window are searched.</summary>
    public bool? ActiveOnly { get; set; }

    /// <summary>When true, off-screen elements are included in candidate collection. Defaults to true.</summary>
    public bool? IncludeOffscreen { get; set; }

    // -------------------------------------------------------------------------
    // Miscellaneous
    // -------------------------------------------------------------------------

    /// <summary>ARIA role or UIA LocalizedControlType hint. Reserved for future matching.</summary>
    public string? Role { get; set; }

    /// <summary>Name hint used for best-match scoring.</summary>
    public string? BestMatch { get; set; }

    // -------------------------------------------------------------------------
    // Rectangle locator filters
    // -------------------------------------------------------------------------
    public int? Left { get; set; }
    public int? Top { get; set; }
    public int? Right { get; set; }
    public int? Bottom { get; set; }
    public int? Width { get; set; }
    public int? Height { get; set; }

    public int? NearX { get; set; }
    public int? NearY { get; set; }
    public int? Tolerance { get; set; }

    /// <summary>When true, limits matches to elements containing the point specified by NearX and NearY.</summary>
    public bool? ContainsPoint { get; set; }

    /// <summary>When true, limits matches to elements intersecting the rectangle specified by Left, Top, Right, and Bottom.</summary>
    public bool? IntersectsRectangle { get; set; }
}
