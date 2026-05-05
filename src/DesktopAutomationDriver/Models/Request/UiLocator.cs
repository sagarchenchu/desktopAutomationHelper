namespace DesktopAutomationDriver.Models.Request;

/// <summary>
/// Identifies a UI element for a /ui endpoint operation.
/// Multiple properties are combined with AND logic.
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

    /// <summary>UIA AutomationId property.</summary>
    public string? AutomationId { get; set; }

    /// <summary>UIA ClassName property.</summary>
    public string? ClassName { get; set; }

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
}
