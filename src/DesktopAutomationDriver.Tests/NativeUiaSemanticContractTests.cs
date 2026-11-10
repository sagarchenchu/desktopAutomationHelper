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
}
