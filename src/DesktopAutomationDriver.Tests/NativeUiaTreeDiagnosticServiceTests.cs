using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaTreeDiagnosticServiceTests
{
    [Fact]
    public void ParseOptions_UsesDumpUiaRequestFields()
    {
        var service = new NativeUiaTreeDiagnosticService(NullLogger<NativeUiaTreeDiagnosticService>.Instance);

        var result = service.DumpTree(
            new UiRequest
            {
                Operation = "dumpuia",
                Root = "activeWindow",
                View = "raw",
                MaxDepth = 10,
                MaxChildren = 500,
                IncludeOffscreen = true,
                NameContains = "DQA",
                IncludePath = true,
                TimeoutMs = 1000,
                Locator = new UiLocator { ControlType = "Button" }
            },
            activeWindowHwnd: null,
            processId: null);

        var json = System.Text.Json.JsonSerializer.Serialize(result);
        Assert.Contains("\"view\":\"raw\"", json);
        Assert.Contains("\"root\":\"activeWindow\"", json);
        Assert.Contains("no-root", json);
    }
}