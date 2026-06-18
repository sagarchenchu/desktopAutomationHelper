using System.Text.Json;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using FlaUI.UIA3;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaBasicOperationHardTimeoutTests
{
    [Fact]
    public void ClickUia_WhenServiceBlocks_ReturnsHardTimeoutJson()
    {
        var basicMock = new Mock<INativeUiaBasicOperationService>();
        basicMock
            .Setup(s => s.Click(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(() =>
            {
                Thread.Sleep(4000);
                return new { operation = "clickuia", success = true };
            });

        using var automation = new UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "basic-hard-timeout-test",
            app,
            automation,
            "UIA3",
            null,
            null);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns(session);

        var service = UiServiceTestFactory.Create(ctxMock.Object, basicMock: basicMock);

        var sw = System.Diagnostics.Stopwatch.StartNew();
        var result = service.Execute(new UiRequest
        {
            Operation = "clickuia",
            TimeoutMs = 1000,
            Locator = new UiLocator { AutomationId = "btnSearch", ControlType = "Button" }
        });
        sw.Stop();

        var json = JsonSerializer.Serialize(result);

        Assert.True(sw.ElapsedMilliseconds < 3000, $"Expected hard timeout around 1s, took {sw.ElapsedMilliseconds}ms");
        Assert.Contains("hard-timeout", json);
        Assert.Contains("\"success\":false", json);
    }
}
