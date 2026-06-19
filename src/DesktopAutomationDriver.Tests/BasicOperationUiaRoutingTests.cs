using System.Text.Json;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class BasicOperationUiaRoutingTests
{
    [Theory]
    [InlineData("clickuia")]
    [InlineData("typeuia")]
    [InlineData("clearuia")]
    [InlineData("focusuia")]
    public void BasicOperationUia_WithoutSession_FailsFastWithoutCallingService(string operation)
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        var request = new UiRequest
        {
            Operation = operation,
            Locator = new UiLocator { AutomationId = "btnSearch", ControlType = "Button" },
            Value = operation == "typeuia" ? "hello" : null,
            TimeoutMs = 3000
        };

        var result = service.Execute(request);
        var json = JsonSerializer.Serialize(result);

        Assert.Contains("no-active-window", json);
        Assert.Contains("\"success\":false", json);
        basicMock.VerifyNoOtherCalls();
    }

    [Fact]
    public void SendKeysUia_WithoutLocator_DoesNotRequireSession()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.SendKeys(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "sendkeysuia", success = true, strategy = "global-sendkeys" });

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        var result = service.Execute(new UiRequest
        {
            Operation = "sendkeysuia",
            Value = "{ENTER}",
            TimeoutMs = 1000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":true", json);
        basicMock.Verify(
            s => s.SendKeys(
                It.Is<UiRequest>(r => r.Operation == "sendkeysuia" && r.Value == "{ENTER}"),
                null,
                null,
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void ClickUia_WithSession_CallsBasicOperationService()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.Click(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "clickuia", success = true, strategy = "invoke-pattern" });

        using var automation = new FlaUI.UIA3.UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "basic-op-test",
            app,
            automation,
            "UIA3",
            null,
            null);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns(session);
        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        var result = service.Execute(new UiRequest
        {
            Operation = "clickuia",
            Locator = new UiLocator { AutomationId = "btnSearch", ControlType = "Button" },
            TimeoutMs = 5000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":true", json);
        basicMock.Verify(
            s => s.Click(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void ClickMenuUia_WithSession_CallsClickMenuOnBasicOperationService()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.ClickMenu(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "clickmenuuia", success = true, strategy = "expandcollapse-expand" });

        using var automation = new FlaUI.UIA3.UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "clickmenu-uia-test",
            app,
            automation,
            "UIA3",
            null,
            null);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns(session);
        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        var result = service.Execute(new UiRequest
        {
            Operation = "clickmenuuia",
            Locator = new UiLocator { Name = "DQA", ControlType = "Menu", MatchMode = "exact" },
            TimeoutMs = 5000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":true", json);
        basicMock.Verify(
            s => s.ClickMenu(
                It.Is<UiRequest>(r => r.Operation == "clickmenuuia"),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }
}
