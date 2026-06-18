using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaComboBoxFindTests
{
    [Fact]
    public void FindComboBox_WithoutSearchContext_FailsFastWithJson()
    {
        var service = new NativeUiaComboBoxService(NullLogger<NativeUiaComboBoxService>.Instance);
        var request = new UiRequest
        {
            Operation = "findcomboboxuia",
            Locator = new UiLocator { AutomationId = "cmbinbound", ControlType = "ComboBox" },
            TimeoutMs = 8000
        };

        var result = service.FindComboBox(request, activeWindowHwnd: null, processId: null);

        var json = System.Text.Json.JsonSerializer.Serialize(result);
        Assert.Contains("\"found\":false", json);
        Assert.Contains("\"success\":false", json);
        Assert.Contains("no-active-window", json);
    }

    [Fact]
    public void SelectComboBoxUia_WithoutSearchContext_FailsFastWithJson()
    {
        var service = new NativeUiaComboBoxService(NullLogger<NativeUiaComboBoxService>.Instance);
        var request = new UiRequest
        {
            Operation = "selectcomboboxuia",
            Locator = new UiLocator { AutomationId = "cmbinbound", ControlType = "ComboBox" },
            Value = "Between",
            TimeoutMs = 15000
        };

        var result = service.SelectComboBox(request, activeWindowHwnd: null, processId: null);

        var json = System.Text.Json.JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":false", json);
        Assert.Contains("no-search-context", json);
    }

    [Theory]
    [InlineData(null, 8000)]
    [InlineData(20000, 15000)]
    public void FindComboBox_DefaultAndMaxTimeout_Clamped(int? requestTimeout, int expectedTimeout)
    {
        var service = new NativeUiaComboBoxService(NullLogger<NativeUiaComboBoxService>.Instance);
        var request = new UiRequest
        {
            Operation = "findcomboboxuia",
            Locator = new UiLocator { AutomationId = "cmbinbound", ControlType = "ComboBox" },
            TimeoutMs = requestTimeout
        };

        var result = service.FindComboBox(request, activeWindowHwnd: null, processId: null);
        var json = System.Text.Json.JsonSerializer.Serialize(result);

        Assert.Contains("no-active-window", json);
        Assert.Equal(expectedTimeout, NativeUiaTimeoutPolicy.Resolve(requestTimeout));
    }
}
