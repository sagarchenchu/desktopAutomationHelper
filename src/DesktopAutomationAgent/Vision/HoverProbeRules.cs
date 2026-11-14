namespace DesktopAutomationAgent.Vision;

/// <summary>
/// Hover is a mouse probe, not a UIA capability. Document outcomes without inventing semantics.
/// </summary>
public static class HoverProbeRules
{
    public static string DescribeOutcome(
        bool structureChanged,
        bool stateChanged,
        string? tooltipText,
        bool noOp)
    {
        if (!string.IsNullOrWhiteSpace(tooltipText))
            return $"hover-tooltip:{tooltipText.Trim()}";
        if (structureChanged)
            return "hover-structure-changed";
        if (stateChanged)
            return "hover-state-changed";
        if (noOp)
            return "hover-noop";
        return "hover-observed";
    }

    public static bool IsUiaCapability => false;
}
