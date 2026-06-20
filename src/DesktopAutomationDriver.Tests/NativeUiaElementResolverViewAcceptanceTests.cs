namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Guards the raw-view element resolver fix required for Menu (50011) clicks such as DQA.
/// </summary>
public class NativeUiaElementResolverViewAcceptanceTests
{
    [Fact]
    public void ElementResolver_UsesViewConditionInBoundedRecursiveSearch()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaElementResolver.cs");

        Assert.Contains("ResolveViewCondition(UiRequest request)", content);
        Assert.Contains("InferDefaultView(UiRequest request)", content);
        Assert.Contains("FindMatchingElementsBoundedRecursive(", content);
        Assert.Contains("IUIAutomationCondition viewCondition", content);
        Assert.Contains("TreeScope.TreeScope_Children,\n                viewCondition)", content);

        var recursiveStart = content.IndexOf(
            "private void FindMatchingElementsBoundedRecursive(",
            StringComparison.Ordinal);
        Assert.True(recursiveStart >= 0);

        var comboRecursiveStart = content.IndexOf(
            "private void FindMatchingComboBoxesBoundedRecursive(",
            StringComparison.Ordinal);
        Assert.True(comboRecursiveStart > recursiveStart);

        var elementRecursiveBlock = content[recursiveStart..comboRecursiveStart];
        Assert.DoesNotContain("_uia.ControlViewCondition", elementRecursiveBlock);

        var comboEnd = content.IndexOf(
            "private static bool ShouldStopAfterFirstStrongMatch",
            StringComparison.Ordinal);
        Assert.True(comboEnd > comboRecursiveStart);

        var comboRecursiveBlock = content[comboRecursiveStart..comboEnd];
        Assert.Contains("viewCondition", comboRecursiveBlock);
        Assert.DoesNotContain("_uia.ControlViewCondition", comboRecursiveBlock);
    }

    [Fact]
    public void ElementResolver_ActiveWindowFallsBackToProcessWindows()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaElementResolver.cs");

        Assert.Contains("stage: \"active-window\"", content);
        Assert.Contains("if (result.Element != null || result.IsAmbiguous)", content);
        Assert.Contains("stage: \"process-window\"", content);
        Assert.Contains("view={viewName}", content);
        Assert.Contains("NativeUiaResolverDiagnostics", content);
        Assert.Contains("ResolveRootMode(UiRequest request)", content);
        Assert.Contains("matched-name-but-rejected", content);
    }

    [Fact]
    public void BasicOperationService_ReturnsResolverDiagnosticsOnNotFound()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaBasicOperationService.cs");

        Assert.Contains("diagnostics = resolveResult.Diagnostics", content);
    }

    [Fact]
    public void UiService_ClickMenuNativeUia_ForcesRawViewAndIncludeOffscreen()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/UiService.cs");

        Assert.Contains("private object? ClickMenuNativeUia(", content);
        Assert.Contains("request.View ??= \"raw\";", content);
        Assert.Contains("request.TreeView ??= \"raw\";", content);
        Assert.Contains("request.IncludeOffscreen ??= true;", content);
    }

    [Fact]
    public void InferDefaultView_MenuLocatorAndClickMenuUia_UseRaw()
    {
        Assert.Equal(
            "raw",
            DesktopAutomationDriver.Services.NativeUia.NativeUiaElementResolver.InferDefaultView(
                new DesktopAutomationDriver.Models.Request.UiRequest
                {
                    Operation = "clickmenuuia",
                    Locator = new DesktopAutomationDriver.Models.Request.UiLocator { Name = "DQA" }
                }));

        Assert.Equal(
            "raw",
            DesktopAutomationDriver.Services.NativeUia.NativeUiaElementResolver.InferDefaultView(
                new DesktopAutomationDriver.Models.Request.UiRequest
                {
                    Operation = "clickuia",
                    Locator = new DesktopAutomationDriver.Models.Request.UiLocator
                    {
                        Name = "DQA",
                        ControlType = "Menu"
                    }
                }));
    }

    [Fact]
    public void InferDefaultView_ComboBoxOperations_DefaultToRaw()
    {
        Assert.Equal(
            "raw",
            DesktopAutomationDriver.Services.NativeUia.NativeUiaElementResolver.InferDefaultView(
                new DesktopAutomationDriver.Models.Request.UiRequest
                {
                    Operation = "findcomboboxuia",
                    Locator = new DesktopAutomationDriver.Models.Request.UiLocator
                    {
                        AutomationId = "cmbTest",
                        ControlType = "ComboBox"
                    }
                }));

        Assert.Equal(
            "raw",
            DesktopAutomationDriver.Services.NativeUia.NativeUiaElementResolver.InferDefaultView(
                new DesktopAutomationDriver.Models.Request.UiRequest
                {
                    Operation = "selectcomboboxuia",
                    Locator = new DesktopAutomationDriver.Models.Request.UiLocator
                    {
                        AutomationId = "cmbTest",
                        ControlType = "ComboBox"
                    }
                }));
    }

    [Fact]
    public void ComboBoxResolver_PassesViewConditionThroughBoundedSearch()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaElementResolver.cs");

        Assert.Contains(
            "var viewCondition = ResolveViewCondition(request);\n        var viewName = ResolveViewName(request);\n        var matches = FindMatchingComboBoxesBounded(\n            root,\n            locator,\n            viewCondition,",
            content);
        Assert.Contains(
            "Native UIA resolver could not find a ComboBox for the locator (view={viewName}).",
            content);
        Assert.Contains(
            "if (operation is \"clickmenuuia\" or \"findcomboboxuia\" or \"selectcomboboxuia\")\n            return \"raw\";",
            content);
    }

    private static string ReadSource(string relativePath)
    {
        var dir = AppContext.BaseDirectory;
        for (var i = 0; i < 8; i++)
        {
            var candidate = Path.GetFullPath(Path.Combine(dir, "..", "..", "..", "..", relativePath));
            if (File.Exists(candidate))
                return File.ReadAllText(candidate).Replace("\r\n", "\n", StringComparison.Ordinal);

            dir = Path.Combine(dir, "..");
        }

        throw new FileNotFoundException($"Could not locate source file: {relativePath}");
    }
}
