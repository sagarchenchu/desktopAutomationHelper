namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Double-click requires policy or known behavior. Do not infer it for every list/grid item.
/// </summary>
public static class DoubleClickPolicy
{
    private static readonly HashSet<string> KnownDoubleClickControlTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "Button",
        "ListItem",
        "TreeItem",
        "DataItem",
        "Hyperlink"
    };

    public static bool IsAllowed(
        string? controlType,
        bool policyAllowsDoubleClick,
        bool knownBehavior = false)
    {
        if (policyAllowsDoubleClick)
            return true;
        if (knownBehavior)
            return true;
        // Without policy/known behavior, only allow when the control type is an explicitly known
        // double-click target AND the caller marked knownBehavior — never infer for arbitrary items.
        _ = controlType;
        _ = KnownDoubleClickControlTypes;
        return false;
    }

    public static string DescribeDenial(string? controlType) =>
        $"Double-click denied for controlType '{controlType ?? "(none)"}' without policy or known behavior.";
}
