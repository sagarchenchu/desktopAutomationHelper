namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Immutable Assistive context captured from the action target <em>before</em> perform().
/// Must be passed through async callbacks; never re-resolved from mutable overlay fields.
/// </summary>
public sealed class AssistiveActionCaptureContext
{
    public RecordedWindowContext? Window { get; init; }

    public string? PageId { get; init; }

    /// <summary>
    /// When true, a next-action BDD is not consumed (e.g. intermediate drag-source selection).
    /// </summary>
    public bool DeferBddConsumption { get; init; }
}
