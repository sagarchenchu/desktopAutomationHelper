namespace DesktopAutomationDriver.Services.ElementResolution;

/// <summary>
/// Options controlling pywinauto-style element resolution behavior.
/// </summary>
public sealed class ResolveOptions
{
    public bool AllowOffscreen { get; init; }
    public bool AllowDisabled { get; init; }
    public bool IncludeHidden { get; init; }

    public bool ReturnCandidates { get; init; } = true;
    public bool ThrowIfNotFound { get; init; }
    public bool ThrowIfAmbiguous { get; init; }

    /// <summary>query | action | click | select | read | scroll | position</summary>
    public string Purpose { get; init; } = "query";
}
