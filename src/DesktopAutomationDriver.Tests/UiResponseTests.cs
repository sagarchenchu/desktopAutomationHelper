using DesktopAutomationDriver.Models.Response;

namespace DesktopAutomationDriver.Tests;

public class UiResponseTests
{
    [Fact]
    public void FromOperationResult_WhenPayloadSuccessFalse_EnvelopeSuccessFalse()
    {
        var payload = new
        {
            operation = "clickmenuuia",
            success = false,
            reason = "element-not-found",
            message = "Native UIA resolver could not find an element for the locator."
        };

        var response = UiResponse.FromOperationResult(payload);

        Assert.False(response.Success);
        Assert.Equal(payload, response.Value);
        Assert.Equal("element-not-found", response.Reason);
        Assert.Equal(payload.message, response.Error);
    }

    [Fact]
    public void FromOperationResult_WhenPayloadSuccessTrue_EnvelopeSuccessTrue()
    {
        var payload = new
        {
            operation = "clickmenuuia",
            success = true,
            strategy = "invoke-pattern"
        };

        var response = UiResponse.FromOperationResult(payload);

        Assert.True(response.Success);
        Assert.Equal(payload, response.Value);
        Assert.Null(response.Error);
    }

    [Fact]
    public void FromOperationResult_WhenPayloadHasNoSuccessFlag_EnvelopeSuccessTrue()
    {
        var payload = new { selected = "Option A" };

        var response = UiResponse.FromOperationResult(payload);

        Assert.True(response.Success);
        Assert.Equal(payload, response.Value);
    }
}
