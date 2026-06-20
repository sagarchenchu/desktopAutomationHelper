using System.Text.Json;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class Phase1QueryUiaRoutingTests
{
    [Theory]
    [InlineData("existsuia")]
    [InlineData("getvalueuia")]
    [InlineData("waituia")]
    public void QueryUia_WithoutSession_FailsFastWithoutCallingService(string operation)
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        var result = service.Execute(new UiRequest
        {
            Operation = operation,
            Locator = new UiLocator { AutomationId = "txtField", ControlType = "Edit" },
            TimeoutMs = 3000
        });

        var json = JsonSerializer.Serialize(result);

        Assert.Contains("no-active-window", json);
        Assert.Contains("\"success\":false", json);
        basicMock.VerifyNoOtherCalls();
    }

    [Fact]
    public void ExistsUia_WithSession_CallsBasicOperationService()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.Exists(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "existsuia", success = true, exists = true, strategy = "element-resolved" });

        using var automation = new FlaUI.UIA3.UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "exists-uia-test",
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
            Operation = "existsuia",
            Locator = new UiLocator { Name = "DQA", ControlType = "Menu", MatchMode = "contains" },
            View = "raw",
            TimeoutMs = 5000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":true", json);
        basicMock.Verify(
            s => s.Exists(
                It.Is<UiRequest>(r => r.Operation == "existsuia" && r.View == "raw"),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void GetValueUia_WithSession_CallsBasicOperationService()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.GetValue(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "getvalueuia", success = true, strategy = "value-pattern", value = "hello" });

        using var automation = new FlaUI.UIA3.UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "getvalue-uia-test",
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
            Operation = "getvalueuia",
            Locator = new UiLocator { AutomationId = "txtField" },
            View = "raw",
            TimeoutMs = 5000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":true", json);
        basicMock.Verify(
            s => s.GetValue(
                It.Is<UiRequest>(r => r.Operation == "getvalueuia"),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void WaitUia_WithSession_CallsBasicOperationService()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.Wait(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "waituia", success = true, strategy = "wait-until-found" });

        using var automation = new FlaUI.UIA3.UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "wait-uia-test",
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
            Operation = "waituia",
            Locator = new UiLocator { Name = "DQA", ControlType = "Menu" },
            View = "raw",
            PollIntervalMs = 200,
            TimeoutMs = 10000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":true", json);
        basicMock.Verify(
            s => s.Wait(
                It.Is<UiRequest>(r =>
                    r.Operation == "waituia"
                    && r.PollIntervalMs == 200
                    && r.TimeoutMs == 10000),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }
}
