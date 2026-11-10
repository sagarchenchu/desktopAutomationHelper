using DesktopAutomationDriver.Services.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaSemanticContractTests
{
    [Fact]
    public void MapCapabilities_FailsClosedForValueUnlessWritable()
    {
        var unknown = NativeUiaSemanticContract.MapCapabilities(["Value"], valueIsReadOnly: null);
        Assert.Equal(["readValue"], unknown);

        var readOnly = NativeUiaSemanticContract.MapCapabilities(["Value"], valueIsReadOnly: true);
        Assert.Equal(["readValue"], readOnly);

        var writable = NativeUiaSemanticContract.MapCapabilities(["Value", "Invoke"], valueIsReadOnly: false);
        Assert.Equal(["invoke", "readValue", "setValue"], writable);
    }

    [Fact]
    public void TruncationWarnings_CoverMaxDepthAndMaxChildren()
    {
        Assert.Null(NativeUiaSemanticContract.MaxChildrenTruncationWarning(availableChildren: 3, maxChildren: 5));
        Assert.Equal(
            "dumpuia hit maxChildren cap (2); 3 sibling(s) omitted.",
            NativeUiaSemanticContract.MaxChildrenTruncationWarning(availableChildren: 5, maxChildren: 2));

        Assert.Contains("maxDepth", NativeUiaSemanticContract.MaxDepthTruncationWarning(4), StringComparison.Ordinal);
    }

    [Fact]
    public void SiblingOrder_SortsByOrdinalThenIdentity()
    {
        var cmp = NativeUiaSemanticContract.CompareSiblingOrder(
            leftOrdinal: 1,
            leftAutomationId: "b",
            leftName: "B",
            leftControlType: "Button",
            leftNodeIndex: 2,
            rightOrdinal: 0,
            rightAutomationId: "a",
            rightName: "A",
            rightControlType: "Button",
            rightNodeIndex: 1);

        Assert.True(cmp > 0);

        var tieBreak = NativeUiaSemanticContract.CompareSiblingOrder(
            0, "a", "A", "Button", 2,
            0, "a", "A", "Button", 1);
        Assert.True(tieBreak > 0);
    }

    [Fact]
    public void SemanticDiagnosticVersion_IsStable()
    {
        Assert.Equal(1, NativeUiaSemanticContract.SemanticDiagnosticVersion);
    }

    [Fact]
    public void TraverseBounded_MaxChildrenKeepsFirstNSiblings()
    {
        // Five children + maxChildren=2 → exactly root + first two children.
        // Informational truncation must not stop sibling visitation of the kept children.
        var root = new SemanticProbeNode
        {
            Id = "root",
            Children =
            [
                new() { Id = "c0" },
                new() { Id = "c1" },
                new() { Id = "c2" },
                new() { Id = "c3" },
                new() { Id = "c4" }
            ]
        };

        var result = NativeUiaSemanticContract.TraverseBounded(
            root,
            maxDepth: 8,
            maxChildren: 2,
            maxNodes: 100);

        Assert.Equal(["root", "c0", "c1"], result.VisitedIds);
        Assert.True(result.IsTruncated);
        Assert.False(result.StopTraversal);
        Assert.Contains(result.Warnings, w => w.Contains("maxChildren", StringComparison.Ordinal));
    }

    [Fact]
    public void TraverseBounded_MaxDepthVisitsAllSiblingsOmitsGrandchild()
    {
        // root→A→A1 and root→B with maxDepth=1 → root, A, B; omit A1.
        // maxDepth truncation must not prevent visiting sibling B after A.
        var root = new SemanticProbeNode
        {
            Id = "root",
            Children =
            [
                new()
                {
                    Id = "A",
                    Children = [new() { Id = "A1" }]
                },
                new() { Id = "B" }
            ]
        };

        var result = NativeUiaSemanticContract.TraverseBounded(
            root,
            maxDepth: 1,
            maxChildren: 50,
            maxNodes: 100);

        Assert.Equal(["root", "A", "B"], result.VisitedIds);
        Assert.DoesNotContain("A1", result.VisitedIds);
        Assert.True(result.IsTruncated);
        Assert.False(result.StopTraversal);
        Assert.Contains(result.Warnings, w => w.Contains("maxDepth", StringComparison.Ordinal));
    }

    [Fact]
    public void TraverseBounded_MaxNodesStopsTraversal()
    {
        var root = new SemanticProbeNode
        {
            Id = "root",
            Children =
            [
                new() { Id = "c0" },
                new() { Id = "c1" },
                new() { Id = "c2" }
            ]
        };

        var result = NativeUiaSemanticContract.TraverseBounded(
            root,
            maxDepth: 8,
            maxChildren: 50,
            maxNodes: 2);

        Assert.Equal(["root", "c0"], result.VisitedIds);
        Assert.True(result.IsTruncated);
        Assert.True(result.StopTraversal);
        Assert.Contains(result.Warnings, w => w.Contains("maxNodes", StringComparison.Ordinal));
    }
}
