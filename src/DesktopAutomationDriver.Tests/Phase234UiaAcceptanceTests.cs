namespace DesktopAutomationDriver.Tests;

public class Phase234UiaAcceptanceTests
{
    [Fact]
    public void UiService_ContainsPhase234OperationMarkers()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/UiService.cs");

        Assert.Contains("\"doubleclickuia\"", content);
        Assert.Contains("\"rightclickuia\"", content);
        Assert.Contains("\"checkuia\"", content);
        Assert.Contains("\"uncheckuia\"", content);
        Assert.Contains("\"selecttabuia\"", content);
        Assert.Contains("\"getgriduia\"", content);
        Assert.Contains("\"selectgridrowuia\"", content);
        Assert.Contains("\"screenshotelementuia\"", content);
        Assert.Contains("DoubleClickNativeUia", content);
        Assert.Contains("GetGridNativeUia", content);
        Assert.Contains("ScreenshotElementNativeUia", content);
        Assert.Contains("ExecuteNativeUiaGridOperation", content);
        Assert.Contains("_nativeUiaGridService.GetGrid", content);
        Assert.Contains("_nativeUiaGridService.SelectGridRow", content);
    }

    [Fact]
    public void BasicOperationService_ContainsPhase2And4Markers()
    {
        var interfaceContent = ReadSource("DesktopAutomationDriver/Services/NativeUia/INativeUiaBasicOperationService.cs");
        var serviceContent = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaBasicOperationService.cs");
        var inputContent = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaInput.cs");

        Assert.Contains("object DoubleClick(", interfaceContent);
        Assert.Contains("object RightClick(", interfaceContent);
        Assert.Contains("object Check(", interfaceContent);
        Assert.Contains("object Uncheck(", interfaceContent);
        Assert.Contains("object SelectTab(", interfaceContent);
        Assert.Contains("object ScreenshotElement(", interfaceContent);

        Assert.Contains("DoubleClickElement", serviceContent);
        Assert.Contains("return (true, \"physical-doubleclick\"", serviceContent);
        Assert.Contains("invoke-twice", serviceContent);
        Assert.DoesNotContain("return (true, \"invoke-pattern\", null, attempted, null)", serviceContent);
        Assert.Contains("RightClickElement", serviceContent);
        Assert.Contains("SetToggleElement", serviceContent);
        Assert.Contains("ExecuteSelectTabOperation", serviceContent);
        Assert.Contains("ExecuteScreenshotElementOperation", serviceContent);

        Assert.Contains("DoubleClickPoint", inputContent);
        Assert.Contains("RightClickPoint", inputContent);
        Assert.Contains("MOUSEEVENTF_RIGHTDOWN", inputContent);
    }

    [Fact]
    public void GridService_ContainsPhase3Markers()
    {
        var interfaceContent = ReadSource("DesktopAutomationDriver/Services/NativeUia/INativeUiaGridService.cs");
        var serviceContent = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaGridService.cs");
        var automationContent = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaAutomation.cs");

        Assert.Contains("object GetGrid(", interfaceContent);
        Assert.Contains("object SelectGridRow(", interfaceContent);
        Assert.Contains("operation = \"getgriduia\"", serviceContent);
        Assert.Contains("operation = \"selectgridrowuia\"", serviceContent);
        Assert.Contains("ReadGridData", serviceContent);
        Assert.Contains("TryGetGridPattern", automationContent);
        Assert.Contains("TryGetTablePattern", automationContent);
    }

    [Fact]
    public void Program_RegistersNativeUiaGridService()
    {
        var content = ReadSource("DesktopAutomationDriver/Program.cs");
        Assert.Contains("INativeUiaGridService, NativeUiaGridService", content);
    }

    private static string ReadSource(string relativePath)
    {
        var path = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..",
            "src",
            relativePath));

        Assert.True(File.Exists(path), $"Expected source file at {path}");
        return File.ReadAllText(path).Replace("\r\n", "\n", StringComparison.Ordinal);
    }
}
