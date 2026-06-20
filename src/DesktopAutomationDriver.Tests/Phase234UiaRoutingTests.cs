using System.Text.Json;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class Phase234UiaRoutingTests
{
    [Theory]
    [InlineData("doubleclickuia")]
    [InlineData("rightclickuia")]
    [InlineData("checkuia")]
    [InlineData("uncheckuia")]
    [InlineData("selecttabuia")]
    [InlineData("screenshotelementuia")]
    public void Phase2Uia_WithoutSession_FailsFastWithoutCallingService(string operation)
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        var result = service.Execute(new UiRequest
        {
            Operation = operation,
            Locator = new UiLocator { AutomationId = "target", ControlType = "Button" },
            Value = operation == "selecttabuia" ? "Tab1" : null,
            TimeoutMs = 3000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("no-active-window", json);
        basicMock.VerifyNoOtherCalls();
    }

    [Theory]
    [InlineData("getgriduia")]
    [InlineData("selectgridrowuia")]
    public void Phase3Uia_WithoutSession_FailsFastWithoutCallingService(string operation)
    {
        var gridMock = new Mock<INativeUiaGridService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = UiServiceTestFactory.Create(ctxMock.Object, gridMock: gridMock);

        var result = service.Execute(new UiRequest
        {
            Operation = operation,
            Locator = new UiLocator { AutomationId = "gridMain", ControlType = "DataGrid" },
            Index = operation == "selectgridrowuia" ? 0 : null,
            TimeoutMs = 3000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("no-active-window", json);
        gridMock.VerifyNoOtherCalls();
    }

    [Fact]
    public void DoubleClickUia_WithSession_CallsBasicOperationService()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.DoubleClick(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "doubleclickuia", success = true, strategy = "physical-doubleclick" });

        using var automation = new FlaUI.UIA3.UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "doubleclick-uia-test",
            app,
            automation,
            "UIA3",
            null,
            null);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns(session);
        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        service.Execute(new UiRequest
        {
            Operation = "doubleclickuia",
            Locator = new UiLocator { AutomationId = "btnSearch", ControlType = "Button" },
            TimeoutMs = 5000
        });

        basicMock.Verify(
            s => s.DoubleClick(
                It.Is<UiRequest>(r => r.Operation == "doubleclickuia"),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void GetGridUia_WithSession_CallsGridService()
    {
        var gridMock = new Mock<INativeUiaGridService>();
        gridMock
            .Setup(s => s.GetGrid(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "getgriduia", success = true, strategy = "grid-pattern", rowCount = 2 });

        using var automation = new FlaUI.UIA3.UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "getgrid-uia-test",
            app,
            automation,
            "UIA3",
            null,
            null);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns(session);
        var service = UiServiceTestFactory.Create(ctxMock.Object, gridMock: gridMock);

        var result = service.Execute(new UiRequest
        {
            Operation = "getgriduia",
            Locator = new UiLocator { AutomationId = "gridMain", ControlType = "DataGrid" },
            TimeoutMs = 5000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("\"success\":true", json);
        gridMock.Verify(
            s => s.GetGrid(
                It.Is<UiRequest>(r => r.Operation == "getgriduia"),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void ScreenshotElementUia_WithSession_CallsBasicOperationService()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.ScreenshotElement(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { operation = "screenshotelementuia", success = true, strategy = "bounding-rect", width = 100, height = 50 });

        using var automation = new FlaUI.UIA3.UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "screenshot-uia-test",
            app,
            automation,
            "UIA3",
            null,
            null);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns(session);
        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        service.Execute(new UiRequest
        {
            Operation = "screenshotelementuia",
            Locator = new UiLocator { AutomationId = "panelMain", ControlType = "Pane" },
            TimeoutMs = 5000
        });

        basicMock.Verify(
            s => s.ScreenshotElement(
                It.Is<UiRequest>(r => r.Operation == "screenshotelementuia"),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }
}
