using DesktopAutomationDriver.Controllers;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Models.Response;
using DesktopAutomationDriver.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

/// <summary>
/// Tests for <see cref="ElementController"/> focusing on session-not-found
/// and error propagation paths (no FlaUI/Windows required).
/// </summary>
public class ElementControllerTests
{
    private readonly Mock<ISessionManager> _sessionManagerMock;
    private readonly Mock<IAutomationService> _automationServiceMock;
    private readonly ElementController _controller;

    public ElementControllerTests()
    {
        _sessionManagerMock = new Mock<ISessionManager>();
        _automationServiceMock = new Mock<IAutomationService>();
        _controller = new ElementController(
            _sessionManagerMock.Object,
            _automationServiceMock.Object,
            NullLogger<ElementController>.Instance);
    }

    // ------------------------------------------------------------------
    // Session-not-found paths (no Windows required)
    // ------------------------------------------------------------------

    [Fact]
    public void FindElement_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.FindElement("bad", new FindElementRequest());
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void FindElements_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.FindElements("bad", new FindElementRequest());
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void Click_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.Click("bad", "el1");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void SendKeys_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.SendKeys("bad", "el1", new SendKeysRequest());
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void GetText_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.GetText("bad", "el1");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void GetAttribute_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.GetAttribute("bad", "el1", "Name");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void IsEnabled_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.IsEnabled("bad", "el1");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void IsDisplayed_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.IsDisplayed("bad", "el1");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void TakeScreenshot_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.TakeScreenshot("bad");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void Clear_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.Clear("bad", "el1");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void DoubleClick_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.DoubleClick("bad", "el1");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    [Fact]
    public void RightClick_WhenSessionNotFound_Returns404()
    {
        SetupSessionNotFound("bad");
        var result = _controller.RightClick("bad", "el1");
        Assert.IsType<NotFoundObjectResult>(result);
    }

    // ------------------------------------------------------------------
    // helpers
    // ------------------------------------------------------------------

    private void SetupSessionNotFound(string sessionId) =>
        _sessionManagerMock.Setup(m => m.GetSession(sessionId)).Returns((AutomationSession?)null);
}
