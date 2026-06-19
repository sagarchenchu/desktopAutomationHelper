using DesktopAutomationDriver.Services.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaTextControlTypeTests
{
    [Theory]
    [InlineData("Menu", 50011)]
    [InlineData("menu", 50011)]
    [InlineData("50011", 50011)]
    [InlineData("ControlType(50011)", 50011)]
    [InlineData("ControlType 50011", 50011)]
    [InlineData("MenuItem", 50010)]
    public void ParseControlTypeId_RecognizesMenuAndAliases(string input, int expected)
    {
        Assert.Equal(expected, NativeUiaText.ParseControlTypeId(input));
    }

    [Fact]
    public void ControlTypeName_ReturnsMenuFor50011()
    {
        Assert.Equal("Menu", NativeUiaText.ControlTypeName(50011));
    }
}
