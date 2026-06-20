namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Guards Phase 1 native UIA query/wait operations wiring in UiService and service implementation.
/// </summary>
public class Phase1QueryUiaAcceptanceTests
{
    [Fact]
    public void UiService_ContainsPhase1QueryOperationMarkers()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/UiService.cs");

        Assert.Contains("\"existsuia\"", content);
        Assert.Contains("\"getvalueuia\"", content);
        Assert.Contains("\"waituia\"", content);
        Assert.Contains("ExistsNativeUia", content);
        Assert.Contains("GetValueNativeUia", content);
        Assert.Contains("WaitNativeUia", content);
        Assert.Contains("ExecuteWaitNativeUiaWithTimeout", content);
        Assert.Contains("_nativeUiaBasicOperationService.Exists", content);
        Assert.Contains("_nativeUiaBasicOperationService.GetValue", content);
        Assert.Contains("_nativeUiaBasicOperationService.Wait", content);
        Assert.Contains("maxHardTimeoutMs: 60000", content);
    }

    [Fact]
    public void BasicOperationService_ContainsPhase1QueryOperationMarkers()
    {
        var interfaceContent = ReadSource("DesktopAutomationDriver/Services/NativeUia/INativeUiaBasicOperationService.cs");
        var serviceContent = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaBasicOperationService.cs");

        Assert.Contains("object Exists(", interfaceContent);
        Assert.Contains("object GetValue(", interfaceContent);
        Assert.Contains("object Wait(", interfaceContent);

        Assert.Contains("ExecuteQueryOperation", serviceContent);
        Assert.Contains("TryResolveLocatedElement", serviceContent);
        Assert.Contains("operation = \"existsuia\"", serviceContent);
        Assert.Contains("operation = \"getvalueuia\"", serviceContent);
        Assert.Contains("operation = \"waituia\"", serviceContent);
        Assert.Contains("strategy = \"wait-until-found\"", serviceContent);
        Assert.Contains("reason = \"wait-timeout\"", serviceContent);
        Assert.Contains("ReadElementValue", serviceContent);
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
