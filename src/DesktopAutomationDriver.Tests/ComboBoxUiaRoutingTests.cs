using System.Text.Json;
using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class ComboBoxUiaRoutingTests
{
    [Fact]
    public void FindComboBoxUia_WithoutSession_FailsFastWithoutCallingService()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var request = new UiRequest
        {
            Operation = "findcomboboxuia",
            Locator = new UiLocator { AutomationId = "cmbinbound", ControlType = "ComboBox" },
            TimeoutMs = 3000
        };

        var result = service.Execute(request);
        var json = JsonSerializer.Serialize(result);

        Assert.Contains("no-active-window", json);
        Assert.Contains("\"success\":false", json);
        Assert.Contains("\"found\":false", json);
        nativeMock.Verify(
            s => s.FindComboBox(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public void FindComboBoxUia_OperationOnly_WithoutSession_FailsFast()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var result = service.Execute(new UiRequest { Operation = "findcomboboxuia", TimeoutMs = 3000 });
        var json = JsonSerializer.Serialize(result);

        Assert.Contains("no-active-window", json);
        nativeMock.VerifyNoOtherCalls();
    }

    [Fact]
    public void SelectComboBoxUia_WithoutSession_FailsFastWithoutCallingService()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var request = new UiRequest
        {
            Operation = "selectcomboboxuia",
            Locator = new UiLocator { AutomationId = "cmbinbound", ControlType = "ComboBox" },
            Value = "Between",
            TimeoutMs = 8000
        };

        var result = service.Execute(request);
        var json = JsonSerializer.Serialize(result);

        Assert.Contains("no-active-window", json);
        nativeMock.Verify(
            s => s.SelectComboBox(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()),
            Times.Never);
    }

    [Fact]
    public void FindComboBoxUia_WithoutSession_WithRequestProcessId_StillFailsFastAtUiService()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var result = service.Execute(new UiRequest
        {
            Operation = "findcomboboxuia",
            Locator = new UiLocator { AutomationId = "missing-combo", ControlType = "ComboBox" },
            ProcessId = 1234,
            TimeoutMs = 3000
        });

        var json = JsonSerializer.Serialize(result);
        Assert.Contains("no-active-window", json);
        nativeMock.VerifyNoOtherCalls();
    }
}
