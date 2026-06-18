using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaBasicOperationServiceTests
{
    [Fact]
    public void Click_WithoutSearchContext_ReturnsElementNotFoundJson()
    {
        var service = new NativeUiaBasicOperationService(NullLogger<NativeUiaBasicOperationService>.Instance);

        var result = service.Click(
            new UiRequest
            {
                Operation = "clickuia",
                Locator = new UiLocator { AutomationId = "missing", ControlType = "Button" },
                TimeoutMs = 1000
            },
            activeWindowHwnd: null,
            processId: null);

        Assert.NotNull(result);
        var json = System.Text.Json.JsonSerializer.Serialize(result);
        Assert.Contains("no-root", json);
        Assert.Contains("\"success\":false", json);
    }

    [Fact]
    public void Type_WithoutValue_ReturnsInvalidRequestJson()
    {
        var service = new NativeUiaBasicOperationService(NullLogger<NativeUiaBasicOperationService>.Instance);

        var result = service.Type(
            new UiRequest
            {
                Operation = "typeuia",
                Locator = new UiLocator { AutomationId = "txtName" },
                TimeoutMs = 1000
            },
            activeWindowHwnd: IntPtr.Zero,
            processId: Environment.ProcessId);

        var json = System.Text.Json.JsonSerializer.Serialize(result);
        Assert.Contains("invalid-request", json);
        Assert.Contains("\"success\":false", json);
    }

    [Fact]
    public void SendKeys_WithoutLocator_AllowsGlobalSendKeys()
    {
        var service = new NativeUiaBasicOperationService(NullLogger<NativeUiaBasicOperationService>.Instance);

        var result = service.SendKeys(
            new UiRequest
            {
                Operation = "sendkeysuia",
                Value = "",
                TimeoutMs = 1000
            },
            activeWindowHwnd: null,
            processId: null);

        var json = System.Text.Json.JsonSerializer.Serialize(result);
        Assert.Contains("invalid-request", json);
    }
}
