using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Middleware;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using DesktopExecutionRuntime;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Moq;

namespace DesktopAutomationDriver.Tests;

public sealed class RuntimeControllerTests
{
    [Fact]
    public void EndpointReturnsObservationsWithoutChangingLegacyStatus()
    {
        var probe = new Mock<IDesktopEnvironmentProbe>();
        probe.Setup(p => p.Capture()).Returns(new DesktopSnapshot { IsWindows = true, SessionId = 0 });
        var result = Assert.IsType<OkObjectResult>(new RuntimeController(probe.Object).GetDiagnostics());
        var response = Assert.IsType<WebDriverResponse<DesktopDiagnostics>>(result.Value);
        Assert.Equal(0, response.Status); // The diagnostic query succeeded, not desktop execution.
        Assert.False(response.Value!.PreflightPassed);
        Assert.Contains("INTERACTIVE_SESSION_UNAVAILABLE", response.Value.BlockingReasons);
        probe.Verify(p => p.Capture(), Times.Once);
        var status = Assert.IsType<OkObjectResult>(new StatusController().GetStatus());
        Assert.True(Assert.IsType<WebDriverResponse<StatusResponse>>(status.Value).Value!.Ready);
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData("Bearer wrong", false)]
    [InlineData("Bearer test-runtime-token", true)]
    public async Task RuntimeDiagnosticsRequiresExistingBearerAuthentication(string? header, bool allowed)
    {
        var context = new DefaultHttpContext();
        context.Request.Path = "/runtime/diagnostics";
        if (header != null) context.Request.Headers.Authorization = header;
        context.Response.Body = new MemoryStream();
        var driverContext = new Mock<IDriverContext>();
        driverContext.Setup(c => c.BearerToken).Returns("test-runtime-token");
        var reachedEndpoint = false;
        var middleware = new BearerTokenMiddleware(_ => { reachedEndpoint = true; return Task.CompletedTask; }, driverContext.Object);
        await middleware.InvokeAsync(context);
        Assert.Equal(allowed, reachedEndpoint);
        Assert.Equal(allowed ? 200 : 401, context.Response.StatusCode);
    }
}
