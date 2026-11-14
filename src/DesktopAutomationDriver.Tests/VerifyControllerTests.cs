using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.Bootstrap;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
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
        mock.Setup(c => c.UserSid).Returns("S-1-5-21-1-2-3-4");
        mock.Setup(c => c.MainPort).Returns(mainPort);
        mock.Setup(c => c.BearerToken).Returns(token);
        mock.Setup(c => c.OperatorToken).Returns(token + "-op");
        mock.Setup(c => c.ProbePortActive).Returns(probePortActive);
        mock.Setup(c => c.ProbePort).Returns(probePort);
        return mock;
    }

    private static VerifyController CreateController(
        IDriverContext context,
        BootstrapDiscoveryMode mode = BootstrapDiscoveryMode.Secure,
        bool acknowledgeLegacyRisk = false)
    {
        var options = Options.Create(new BootstrapOptions
        {
            Mode = mode,
            AcknowledgeLegacyVerifyRisk = acknowledgeLegacyRisk
        });
        return new VerifyController(context, options);
    }

    [Fact]
    public void SecureMode_Returns403WithoutToken()
    {
        const string token = "must-not-appear";
        var controller = CreateController(BuildContextMock(token: token).Object);

        var result = Assert.IsType<ObjectResult>(controller.Verify());
        Assert.Equal(StatusCodes.Status403Forbidden, result.StatusCode);

        var json = System.Text.Json.JsonSerializer.Serialize(result.Value);
        Assert.DoesNotContain(token, json, StringComparison.Ordinal);
        Assert.Contains("secure_bootstrap_required", json, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LegacyMode_Returns200WithToken_WhenRiskAcknowledged()
    {
        var mock = BuildContextMock(token: "abc123token");
        var controller = CreateController(
            mock.Object,
            BootstrapDiscoveryMode.LegacyVerify,
            acknowledgeLegacyRisk: true);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Equal("abc123token", response.Value!.Token);
        Assert.Equal("Bearer abc123token", response.Value.AuthorizationHeader);
    }

    [Fact]
    public void LegacyMode_ResponseContainsMainPortAndUsername()
    {
        var mock = BuildContextMock(username: "domain\\alice", mainPort: 31500, token: "t");
        var controller = CreateController(
            mock.Object,
            BootstrapDiscoveryMode.LegacyVerify,
            acknowledgeLegacyRisk: true);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Equal(31500, response.Value!.Port);
        Assert.Equal("domain\\alice", response.Value.Username);
        Assert.True(response.Value.Running);
    }

    [Fact]
    public void LegacyMode_WhenProbePortActive_IncludesProbePort()
    {
        var mock = BuildContextMock(probePortActive: true, probePort: 9102);
        var controller = CreateController(
            mock.Object,
            BootstrapDiscoveryMode.LegacyVerify,
            acknowledgeLegacyRisk: true);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Equal(9102, response.Value!.ProbePort);
    }

    [Fact]
    public void LegacyMode_WhenProbePortInactive_ProbePortIsNull()
    {
        var mock = BuildContextMock(probePortActive: false);
        var controller = CreateController(
            mock.Object,
            BootstrapDiscoveryMode.LegacyVerify,
            acknowledgeLegacyRisk: true);

        var ok = Assert.IsType<OkObjectResult>(controller.Verify());
        var response = Assert.IsType<WebDriverResponse<VerifyResponse>>(ok.Value);
        Assert.Null(response.Value!.ProbePort);
    }
}
