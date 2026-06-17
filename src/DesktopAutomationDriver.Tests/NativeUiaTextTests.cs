using DesktopAutomationDriver.Services.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaTextTests
{
    [Theory]
    [InlineData("Inbound", "Inbound", true)]
    [InlineData("  Inbound  ", "Inbound", true)]
    [InlineData("In&bound", "Inbound", true)]
    [InlineData("Between", "Inbound", false)]
    public void Matches_ExactMode_ComparesNormalizedValues(string candidate, string requested, bool expected)
    {
        Assert.Equal(expected, NativeUiaText.Matches(candidate, requested, "exact"));
    }

    [Theory]
    [InlineData("Some Very Far Item", "Far", true)]
    [InlineData("Item 1", "Item", true)]
    [InlineData("Alpha", "Beta", false)]
    public void Matches_ContainsMode_Works(string candidate, string requested, bool expected)
    {
        Assert.Equal(expected, NativeUiaText.Matches(candidate, requested, "contains"));
    }

    [Theory]
    [InlineData("Inbound option", "Inbound", true)]
    [InlineData("Between", "Inbound", false)]
    public void ValuesEquivalent_AllowsContainsFallback(string actual, string requested, bool expected)
    {
        Assert.Equal(expected, NativeUiaText.ValuesEquivalent(actual, requested));
    }

    [Fact]
    public void Normalize_RemovesMnemonicAndCollapsesWhitespace()
    {
        var normalized = NativeUiaText.Normalize("  In&bound   option ");
        Assert.Equal("Inbound option", normalized);
    }

    [Theory]
    [InlineData("ComboBox", 50003)]
    [InlineData("Edit", 50004)]
    [InlineData("ListItem", 50007)]
    public void ParseControlTypeId_MapsKnownNames(string controlType, int expected)
    {
        Assert.Equal(expected, NativeUiaText.ParseControlTypeId(controlType));
    }
}
