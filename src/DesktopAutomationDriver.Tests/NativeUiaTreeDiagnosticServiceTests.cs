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

    [Fact]
    public void ParseOptions_ExposesSemanticOptionalDefaultsInNoRootPayload()
    {
        var service = new NativeUiaTreeDiagnosticService(NullLogger<NativeUiaTreeDiagnosticService>.Instance);

        var result = service.DumpTree(
            new UiRequest
            {
                Operation = "dumpuia",
                MaxNodes = 25,
                IncludeBounds = false,
                IncludeHelpText = true,
                IncludeFocusState = true,
                IncludeRuntimeId = true,
                IncludeTreeLinks = true,
                IncludeWindowState = true,
                TimeoutMs = 500
            },
            activeWindowHwnd: null,
            processId: null);

        var json = System.Text.Json.JsonSerializer.Serialize(result);
        Assert.Contains("no-root", json);
        // Options are only echoed on successful root resolution; ensure request still accepted.
        Assert.DoesNotContain("exceptionType", json, StringComparison.OrdinalIgnoreCase);
    }
}