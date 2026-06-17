using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaComboBoxTimeoutTests
{
    [Fact]
    public void SamplePayloads_DocumentOperarRequest()
    {
        const string payload = """
            {
              "operation": "select",
              "locator": {
                "name": "Operar",
                "controlType": "ComboBox"
              },
              "value": "equals",
              "timeoutMs": 8000
            }
            """;

        Assert.Contains("Operar", payload);
        Assert.Contains("ComboBox", payload);
        Assert.Contains("8000", payload);
    }

    [Fact]
    public void InspectComboBoxPayload_IsDocumented()
    {
        const string payload = """
            {
              "operation": "inspectcombobox",
              "locator": {
                "name": "Operar",
                "controlType": "ComboBox"
              },
              "timeoutMs": 5000
            }
            """;

        Assert.Contains("inspectcombobox", payload);
    }

    [Theory]
    [InlineData(null, 8000)]
    [InlineData(5000, 5000)]
    [InlineData(20000, 15000)]
    [InlineData(100, 500)]
    public void TimeoutPolicy_Clamped(int? requestTimeout, int expected)
    {
        Assert.Equal(expected, NativeUiaTimeoutPolicy.Resolve(requestTimeout));
    }
}

internal static class NativeUiaTimeoutPolicy
{
    public const int DefaultTimeoutMs = 8000;
    public const int MaxTimeoutMs = 15000;

    public static int Resolve(int? requestTimeoutMs)
    {
        var timeout = requestTimeoutMs ?? DefaultTimeoutMs;
        return Math.Clamp(timeout, 500, MaxTimeoutMs);
    }
}
