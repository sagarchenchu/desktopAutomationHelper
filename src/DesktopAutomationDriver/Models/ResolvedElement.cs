using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Models;

/// <summary>
/// Carries the outcome of a central element resolution attempt.
/// </summary>
public class ResolvedElement
{
    /// <summary>The resolved element, or null when not found.</summary>
    public AutomationElement? Element { get; set; }

    /// <summary>The search strategy that produced the result.</summary>
    public string Strategy { get; set; } = string.Empty;

    /// <summary>The root element strategy used for this search.</summary>
    public string RootStrategy { get; set; } = string.Empty;

    /// <summary>Diagnostics populated on failure, ambiguity, or when requested.</summary>
    public ElementSearchDiagnostics? Diagnostics { get; set; }
}
