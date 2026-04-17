using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;

namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Tests for the request/response model classes.
/// </summary>
public class ModelTests
{
    [Fact]
    public void CreateSessionRequest_DefaultsAreCorrect()
    {
        var request = new CreateSessionRequest();
        Assert.NotNull(request.DesiredCapabilities);
    }

    [Fact]
    public void DesiredCapabilities_DefaultUiaTypeIsUIA3()
    {
        var caps = new DesiredCapabilities();
        Assert.Equal("UIA3", caps.UiaType);
    }

    [Fact]
    public void DesiredCapabilities_DefaultLaunchDelayIs1000()
    {
        var caps = new DesiredCapabilities();
        Assert.Equal(1000, caps.LaunchDelay);
    }

    [Fact]
    public void FindElementRequest_DefaultValuesAreEmpty()
    {
        var req = new FindElementRequest();
        Assert.Equal(string.Empty, req.Using);
        Assert.Equal(string.Empty, req.Value);
    }

    [Fact]
    public void SendKeysRequest_DefaultValueIsEmptyArray()
    {
        var req = new SendKeysRequest();
        Assert.NotNull(req.Value);
        Assert.Empty(req.Value);
    }

    [Fact]
    public void WebDriverResponse_SuccessReturnsStatusZero()
    {
        var response = WebDriverResponse<string>.Success("hello", "session1");
        Assert.Equal(0, response.Status);
        Assert.Equal("hello", response.Value);
        Assert.Equal("session1", response.SessionId);
    }

    [Fact]
    public void WebDriverResponse_ErrorReturnsNonZeroStatus()
    {
        var response = WebDriverResponse<ErrorDetail>.Error(7, "Not found", "no such element");
        Assert.Equal(7, response.Status);
        Assert.NotNull(response.Value);
        Assert.Equal("Not found", response.Value!.Message);
        Assert.Equal("no such element", response.Value!.Error);
    }

    [Fact]
    public void SessionResponse_DefaultsAreCorrect()
    {
        var resp = new SessionResponse();
        Assert.Equal(string.Empty, resp.SessionId);
        Assert.NotNull(resp.Capabilities);
    }

    [Fact]
    public void StatusResponse_DefaultReadyIsTrue()
    {
        var resp = new StatusResponse();
        Assert.True(resp.Ready);
        Assert.NotNull(resp.Build);
        Assert.Equal("1.0.0", resp.Build.Version);
    }

    [Fact]
    public void ElementResponse_DefaultElementIdIsEmpty()
    {
        var resp = new ElementResponse();
        Assert.Equal(string.Empty, resp.ElementId);
    }

    [Fact]
    public void VerifyResponse_DefaultRunningIsTrue()
    {
        var resp = new VerifyResponse();
        Assert.True(resp.Running);
    }

    [Fact]
    public void VerifyResponse_DefaultsAreEmptyOrZero()
    {
        var resp = new VerifyResponse();
        Assert.Equal(string.Empty, resp.Username);
        Assert.Equal(0, resp.Port);
        Assert.Null(resp.ProbePort);
        Assert.Equal(string.Empty, resp.Token);
        Assert.Equal(string.Empty, resp.AuthorizationHeader);
    }

    [Fact]
    public void UiLocator_XPathPropertyDefaultsToNull()
    {
        var locator = new UiLocator();
        Assert.Null(locator.XPath);
    }

    [Fact]
    public void UiLocator_XPathPropertyCanBeSet()
    {
        var locator = new UiLocator { XPath = "//Button[@Name='Save']" };
        Assert.Equal("//Button[@Name='Save']", locator.XPath);
    }
}
