using System.Text.Json;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using FlaUI.UIA3;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaHardTimeoutTests
{
    [Fact]
    public void FindComboBoxUia_WhenServiceBlocks_ReturnsHardTimeoutJson()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>();
        nativeMock
            .Setup(s => s.FindComboBox(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(() =>
            {
                Thread.Sleep(4000);
                return new { operation = "findcomboboxuia", success = true, found = true };
            });

        using var automation = new UIA3Automation();
        var app = FlaUI.Core.Application.Attach(Environment.ProcessId);
        using var session = new AutomationSession(
            "hard-timeout-test",
            app,
            automation,
            "UIA3",
            null,
            null);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns(session);

        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var sw = System.Diagnostics.Stopwatch.StartNew();
        var result = service.Execute(new UiRequest
        {
            Operation = "findcomboboxuia",
            TimeoutMs = 1000,
            Locator = new UiLocator { AutomationId = "cmbText", ControlType = "ComboBox" }
        });
        sw.Stop();

        var json = JsonSerializer.Serialize(result);

        Assert.True(sw.ElapsedMilliseconds < 3000, $"Expected hard timeout around 1s, took {sw.ElapsedMilliseconds}ms");
        Assert.Contains("hard-timeout", json);
        Assert.Contains("\"success\":false", json);
        Assert.Contains("\"found\":false", json);
    }
}
