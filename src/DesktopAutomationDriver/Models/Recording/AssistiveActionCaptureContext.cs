namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// Pre-captured Assistive context used when enriching a successfully recorded action.
/// Built before action execution so window changes cannot corrupt page association.
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
