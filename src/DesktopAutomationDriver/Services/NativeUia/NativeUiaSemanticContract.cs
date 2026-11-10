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

    /// <summary>
    /// Pure bounded DFS used by unit tests and as the behavioral contract for dumpuia.
    /// <see cref="SemanticBoundedTraversalResult.IsTruncated"/> is informational
    /// (maxDepth / maxChildren / maxNodes). <see cref="SemanticBoundedTraversalResult.StopTraversal"/>
    /// is set only for maxNodes (timeout/cancellation are handled by the live collector).
    /// </summary>
    public static SemanticBoundedTraversalResult TraverseBounded(
        SemanticProbeNode root,
        int maxDepth,
        int maxChildren,
        int maxNodes)
    {
        ArgumentNullException.ThrowIfNull(root);
        if (maxDepth < 0)
            throw new ArgumentOutOfRangeException(nameof(maxDepth));
        if (maxChildren < 0)
            throw new ArgumentOutOfRangeException(nameof(maxChildren));
        if (maxNodes < 1)
            throw new ArgumentOutOfRangeException(nameof(maxNodes));

        var visited = new List<string>();
        var warnings = new List<string>();
        var isTruncated = false;
        var stopTraversal = false;

        void Visit(SemanticProbeNode node, int depth)
        {
            if (stopTraversal)
                return;

            if (visited.Count >= maxNodes)
            {
                isTruncated = true;
                stopTraversal = true;
                warnings.Add("dumpuia reached requested maxNodes cap.");
                return;
            }

            visited.Add(node.Id);

            if (stopTraversal)
                return;

            if (depth >= maxDepth)
            {
                if (node.Children.Count > 0)
                {
                    isTruncated = true;
                    var depthWarning = MaxDepthTruncationWarning(maxDepth);
                    if (!warnings.Contains(depthWarning, StringComparer.Ordinal))
                        warnings.Add(depthWarning);
                }

                return;
            }

            var childrenWarning = MaxChildrenTruncationWarning(node.Children.Count, maxChildren);
            if (childrenWarning is not null)
            {
                isTruncated = true;
                if (!warnings.Contains(childrenWarning, StringComparer.Ordinal))
                    warnings.Add(childrenWarning);
            }

            var childCount = Math.Min(node.Children.Count, maxChildren);
            for (var i = 0; i < childCount; i++)
            {
                Visit(node.Children[i], depth + 1);
                if (stopTraversal)
                    return;
            }
        }

        Visit(root, depth: 0);

        return new SemanticBoundedTraversalResult
        {
            VisitedIds = visited,
            IsTruncated = isTruncated,
            StopTraversal = stopTraversal,
            Warnings = warnings
        };
    }
}

/// <summary>
/// In-memory probe node for bounded traversal contract tests.
/// </summary>
internal sealed class SemanticProbeNode
{
    public required string Id { get; init; }
    public IReadOnlyList<SemanticProbeNode> Children { get; init; } = Array.Empty<SemanticProbeNode>();
}

/// <summary>
/// Result of <see cref="NativeUiaSemanticContract.TraverseBounded"/>.
/// </summary>
internal sealed class SemanticBoundedTraversalResult
{
    public required IReadOnlyList<string> VisitedIds { get; init; }
    public required bool IsTruncated { get; init; }
    public required bool StopTraversal { get; init; }
    public required IReadOnlyList<string> Warnings { get; init; }
}
