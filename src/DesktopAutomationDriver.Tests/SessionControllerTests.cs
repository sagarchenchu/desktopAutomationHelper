using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Tests for <see cref="SessionController"/>.
/// All FlaUI / OS interactions are mocked via <see cref="ISessionManager"/>.
/// </summary>
public class SessionControllerTests
{
    private readonly Mock<ISessionManager> _sessionManagerMock;
    private readonly SessionController _controller;

    public SessionControllerTests()
    {
        _sessionManagerMock = new Mock<ISessionManager>();
        _controller = new SessionController(
            _sessionManagerMock.Object,
            NullLogger<SessionController>.Instance);
    }

    [Fact]
    public void GetSession_WhenNotFound_Returns404()
    {
        _sessionManagerMock.Setup(m => m.GetSession("missing")).Returns((AutomationSession?)null);

        var result = _controller.GetSession("missing");

        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void DeleteSession_WhenNotFound_Returns404()
    {
        _sessionManagerMock.Setup(m => m.GetSession("missing")).Returns((AutomationSession?)null);

        var result = _controller.DeleteSession("missing");

        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void ListSessions_ReturnsAllSessionIds()
    {
        var ids = new List<string> { "id1", "id2" };
        _sessionManagerMock.Setup(m => m.GetAllSessionIds()).Returns(ids);

        var result = _controller.ListSessions();

        var ok = Assert.IsType<OkObjectResult>(result);
        Assert.NotNull(ok.Value);
    }

    [Fact]
    public void CreateSession_WhenArgumentExceptionThrown_ReturnsBadRequest()
    {
        _sessionManagerMock
            .Setup(m => m.CreateSession(It.IsAny<DesiredCapabilities>()))
            .Throws(new ArgumentException("Either 'App' or 'AppName' must be specified."));

        var request = new CreateSessionRequest
        {
            DesiredCapabilities = new DesiredCapabilities()
        };

        var result = _controller.CreateSession(request);

        Assert.IsType<BadRequestObjectResult>(result);
    }

    [Fact]
    public void CreateSession_WhenInvalidOperationThrown_Returns422()
    {
        _sessionManagerMock
            .Setup(m => m.CreateSession(It.IsAny<DesiredCapabilities>()))
            .Throws(new InvalidOperationException("No running process found."));

        var request = new CreateSessionRequest
        {
            DesiredCapabilities = new DesiredCapabilities { AppName = "nonexistent" }
        };

        var result = _controller.CreateSession(request);

        Assert.IsType<UnprocessableEntityObjectResult>(result);
    }

    [Fact]
    public void CreateSession_WhenUnexpectedExceptionThrown_Returns500()
    {
        _sessionManagerMock
            .Setup(m => m.CreateSession(It.IsAny<DesiredCapabilities>()))
            .Throws(new Exception("Unexpected failure"));

        var request = new CreateSessionRequest
        {
            DesiredCapabilities = new DesiredCapabilities { App = "notepad.exe" }
        };

        var result = _controller.CreateSession(request);

        var statusResult = Assert.IsType<ObjectResult>(result);
        Assert.Equal(500, statusResult.StatusCode);
    }
}
