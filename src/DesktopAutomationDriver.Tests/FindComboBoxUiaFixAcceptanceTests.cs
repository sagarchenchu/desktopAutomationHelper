namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Guards the findcomboboxuia hard-timeout fix on branch sagar/native-uia-combobox-selection.
/// These tests fail if the cooperative-only wrapper or old resolver is reintroduced.
/// </summary>
public class FindComboBoxUiaFixAcceptanceTests
{
    [Fact]
    public void UiService_ContainsHardTimeoutWrapperMarkers()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/UiService.cs");

        Assert.Contains("ExecuteNativeUiaWithTimeout", content);
        Assert.Contains("Task.Run<object?>", content);
        Assert.Contains("Task.WhenAny", content);
        Assert.Contains("hard-timeout", content);
        Assert.Contains("FindComboBoxNativeUia", content);
        Assert.Contains("_nativeUiaComboBoxService.FindComboBox", content);
        Assert.DoesNotContain("NativeUiaComboBoxSelector", content);
        Assert.DoesNotContain("ExecuteNativeUiaComboBoxOperation", content);
    }

    [Fact]
    public void Resolver_ContainsEarlyMatchTraversalMarkers()
    {
        var content = ReadSource("DesktopAutomationDriver/Services/NativeUia/NativeUiaElementResolver.cs");

        Assert.Contains("FindMatchingComboBoxesBounded", content);
        Assert.Contains("ShouldStopAfterFirstStrongMatch", content);
        Assert.Contains("FindMatchingComboBoxesBoundedRecursive(", content);
        Assert.Contains("IUIAutomationCondition viewCondition", content);
        Assert.DoesNotContain("FindComboBoxesBounded", content);
        Assert.DoesNotContain("TreeScope_Subtree", content);
        Assert.DoesNotContain("TreeScope.Subtree", content);
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
