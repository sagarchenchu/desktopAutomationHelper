namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// Shared Phase 4A semantic diagnostic contract helpers for dumpuia.
/// </summary>
internal static class NativeUiaSemanticContract
{
    public const int SemanticDiagnosticVersion = 1;

    public static List<string> MapCapabilities(
        IReadOnlyList<string> patterns,
        bool? valueIsReadOnly)
    {
        var capabilities = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var pattern in patterns)
        {
            switch (pattern)
            {
                case "Invoke":
                    capabilities.Add("invoke");
                    break;
                case "ExpandCollapse":
                    capabilities.Add("expandCollapse");
                    break;
                case "SelectionItem":
                    capabilities.Add("selectItem");
                    break;
                case "Selection":
                    capabilities.Add("selectionContainer");
                    break;
                case "Toggle":
                    capabilities.Add("toggle");
                    break;
                case "Value":
                    // Fail closed: expose setValue only when CurrentIsReadOnly is explicitly false.
                    capabilities.Add("readValue");
                    if (valueIsReadOnly == false)
                        capabilities.Add("setValue");
                    break;
                case "Scroll":
                    capabilities.Add("scroll");
                    break;
                case "Grid":
                    capabilities.Add("grid");
                    break;
                case "Window":
                    capabilities.Add("window");
                    break;
            }
        }

        return capabilities.ToList();
    }

    public static string? MaxChildrenTruncationWarning(int availableChildren, int maxChildren)
    {
        if (availableChildren <= maxChildren)
            return null;

        var omitted = availableChildren - maxChildren;
        return $"dumpuia hit maxChildren cap ({maxChildren}); {omitted} sibling(s) omitted.";
    }

    public static string MaxDepthTruncationWarning(int maxDepth) =>
        $"dumpuia hit maxDepth cap ({maxDepth}); deeper descendants omitted.";

    public static int CompareSiblingOrder(
        int leftOrdinal,
        string? leftAutomationId,
        string? leftName,
        string? leftControlType,
        int leftNodeIndex,
        int rightOrdinal,
        string? rightAutomationId,
        string? rightName,
        string? rightControlType,
        int rightNodeIndex)
    {
        var cmp = leftOrdinal.CompareTo(rightOrdinal);
        if (cmp != 0)
            return cmp;

        cmp = string.Compare(leftAutomationId, rightAutomationId, StringComparison.Ordinal);
        if (cmp != 0)
            return cmp;

        cmp = string.Compare(leftName, rightName, StringComparison.Ordinal);
        if (cmp != 0)
            return cmp;

        cmp = string.Compare(leftControlType, rightControlType, StringComparison.Ordinal);
        if (cmp != 0)
            return cmp;

        return leftNodeIndex.CompareTo(rightNodeIndex);
    }
}
