using DesktopAutomationDriver.Models.Request;
using FlaUI.Core.AutomationElements;

namespace DesktopAutomationDriver.Services;

/// <summary>
/// Carries the outcome of a central element resolution attempt.
/// Returned by <c>UiService.ResolveElement</c>.
/// </summary>
internal sealed class ElementResolveResult
{
    /// <summary>The resolved element, or <see langword="null"/> when not found.</summary>
    public AutomationElement? Element { get; init; }

    /// <summary><c>true</c> when <see cref="Element"/> is non-null.</summary>
    public bool Found => Element != null;

    /// <summary>The search strategy that produced the result (e.g. "automationid-controltype", "scored-unique").</summary>
    public string Strategy { get; init; } = string.Empty;

    /// <summary>The root element strategy used for this search (e.g. "active-window", "desktop", "parent-locator").</summary>
    public string RootStrategy { get; init; } = string.Empty;

    /// <summary>Candidate diagnostics populated on failure or when diagnostics are requested.</summary>
    public List<ElementCandidateDto> Candidates { get; init; } = new();

    /// <summary>Human-readable error or diagnostic messages.</summary>
    public List<string> Errors { get; init; } = new();

    /// <summary>The locator that was used for this search.</summary>
    public UiLocator? Locator { get; init; }

    /// <summary>Number of candidates collected (before filtering/scoring).</summary>
    public int CandidateCount { get; init; }

    /// <summary><c>true</c> when more than one candidate passed filtering (ambiguous match).</summary>
    public bool Ambiguous { get; init; }
}

/// <summary>
/// Snapshot of a candidate element captured during element resolution for diagnostics.
/// </summary>
internal sealed class ElementCandidateDto
{
    public int Index { get; init; }
    public string Name { get; init; } = string.Empty;
    public string AutomationId { get; init; } = string.Empty;
    public string ClassName { get; init; } = string.Empty;
    public string ControlType { get; init; } = string.Empty;
    public string FrameworkId { get; init; } = string.Empty;
    public int? ProcessId { get; init; }
    public long? Hwnd { get; init; }
    public string Value { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
    public bool? IsEnabled { get; init; }
    public bool? IsOffscreen { get; init; }
    public object? Rectangle { get; init; }
    public int Score { get; init; }
    public string Reason { get; init; } = string.Empty;
}
