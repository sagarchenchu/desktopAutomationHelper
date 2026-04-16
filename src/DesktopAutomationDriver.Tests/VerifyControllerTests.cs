using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;
using Moq;

namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Tests for <see cref="VerifyController"/>.
/// </summary>
public class VerifyControllerTests
{
    private static Mock<IDriverContext> BuildContextMock(
        string username = "testuser",
        int mainPort = 32000,
        string token = "mytoken",
        bool probePortActive = true,
        int probePort = 9102)
    {
        var mock = new Mock<IDriverContext>();
        mock.Setup(c => c.Username).Returns(username);
        mock.Setup(c => c.MainPort).Returns(mainPort);
        mock.Setup(c => c.BearerToken).Returns(token);
        mock.Setup(c => c.ProbePortActive).Returns(probePortActive);
        mock.Setup(c => c.ProbePort).Returns(probePort);
        return mock;
    }

    [Fact]
    public void Verify_Returns200()
    {
        var controller = new VerifyController(BuildContextMock().Object);
        var result = controller.Verify();
        Assert.IsType<OkObjectResult>(result);
    }

    [Fact]
    public void Verify_ResponseContainsToken()
    {
        var mock = BuildContextMock(token: "abc123token");
        var controller = new VerifyController(mock.Object);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Equal("abc123token", response.Value!.Token);
    }

    [Fact]
    public void Verify_AuthorizationHeaderFormatIsCorrect()
    {
        var mock = BuildContextMock(token: "mytoken123");
        var controller = new VerifyController(mock.Object);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Equal("Bearer mytoken123", response.Value!.AuthorizationHeader);
    }

    [Fact]
    public void Verify_ResponseContainsMainPort()
    {
        var mock = BuildContextMock(mainPort: 31500);
        var controller = new VerifyController(mock.Object);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Equal(31500, response.Value!.Port);
    }

    [Fact]
    public void Verify_ResponseContainsUsername()
    {
        var mock = BuildContextMock(username: "domain\\alice");
        var controller = new VerifyController(mock.Object);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Equal("domain\\alice", response.Value!.Username);
    }

    [Fact]
    public void Verify_WhenProbePortActive_IncludesProbePort()
    {
        var mock = BuildContextMock(probePortActive: true, probePort: 9102);
        var controller = new VerifyController(mock.Object);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Equal(9102, response.Value!.ProbePort);
    }

    [Fact]
    public void Verify_WhenProbePortInactive_ProbePortIsNull()
    {
        var mock = BuildContextMock(probePortActive: false);
        var controller = new VerifyController(mock.Object);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Null(response.Value!.ProbePort);
    }

    [Fact]
    public void Verify_RunningIsAlwaysTrue()
    {
        var controller = new VerifyController(BuildContextMock().Object);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.True(response.Value!.Running);
    }
}
