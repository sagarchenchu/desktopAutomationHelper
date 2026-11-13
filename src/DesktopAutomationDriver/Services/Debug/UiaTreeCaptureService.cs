namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Captures bounded UIA ancestry/children/patterns for Debug Mode.
/// Delegates live probing to <see cref="IUiaElementProbe"/>; pure helpers live here.
/// </summary>
public sealed class UiaTreeCaptureService
{
    public const int DefaultMaxAncestry = 8;
    public const int DefaultMaxChildren = 25;
    public const int AbsoluteMaxAncestry = 20;
    public const int AbsoluteMaxChildren = 100;

    private readonly IUiaElementProbe _probe;

    public UiaTreeCaptureService(IUiaElementProbe probe)
    {
        _probe = probe ?? throw new ArgumentNullException(nameof(probe));
    }

    public DebugTreeCapture? CaptureUnderCursor(
        Point screenPoint,
        int? maxAncestry = null,
        int? maxChildren = null)
    {
        var ancestry = Clamp(maxAncestry ?? DefaultMaxAncestry, 0, AbsoluteMaxAncestry);
        var children = Clamp(maxChildren ?? DefaultMaxChildren, 1, AbsoluteMaxChildren);
        return _probe.CaptureAtPoint(screenPoint, ancestry, children);
    }

    /// <summary>
    /// Builds a capture from already-materialized nodes (unit tests / fixtures).
    /// </summary>
    public static DebugTreeCapture FromNodes(
        DebugUiaNode target,
        IReadOnlyList<DebugUiaNode>? ancestry = null,
        IReadOnlyList<DebugUiaNode>? children = null,
        DebugWindowIdentity? window = null,
        Point? cursor = null)
    {
        ArgumentNullException.ThrowIfNull(target);
        var limitedAncestry = (ancestry ?? []).Take(AbsoluteMaxAncestry).ToList();
        var limitedChildren = (children ?? []).Take(AbsoluteMaxChildren).ToList();
        return new DebugTreeCapture
        {
            Target = target,
            Ancestry = limitedAncestry,
            Children = limitedChildren,
            Window = window,
            Cursor = cursor ?? Point.Empty,
            CapturedAtUtc = DateTimeOffset.UtcNow
        };
    }

    public static bool IsSemanticallyInaccessible(DebugUiaNode node)
    {
        ArgumentNullException.ThrowIfNull(node);
        if (!node.IsSemanticAccessible)
            return true;
        if (!string.IsNullOrWhiteSpace(node.InaccessibilityReason))
            return true;
        // Custom-drawn surfaces often report as Custom with empty identity.
        if (string.Equals(node.ControlType, "Custom", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(node.AutomationId)
            && string.IsNullOrWhiteSpace(node.Name))
            return true;
        return false;
    }

    private static int Clamp(int value, int min, int max) =>
        value < min ? min : value > max ? max : value;
}
