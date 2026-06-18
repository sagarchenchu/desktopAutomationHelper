using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class ComboBoxUiaRoutingTests
{
    [Fact]
    public void FindComboBoxUia_WithoutSession_RoutesToNativeServiceWithNullContext()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);
        nativeMock
            .Setup(s => s.FindComboBox(
                It.Is<UiRequest>(r => r.Operation == "findcomboboxuia"),
                null,
                null,
                It.IsAny<CancellationToken>()))
            .Returns(new
            {
                operation = "findcomboboxuia",
                found = false,
                success = false,
                stage = "no-search-context"
            });

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var request = new UiRequest
        {
            Operation = "findcomboboxuia",
            Locator = new UiLocator { AutomationId = "cmbinbound", ControlType = "ComboBox" },
            TimeoutMs = 8000
        };

        var result = service.Execute(request);

        Assert.NotNull(result);
        nativeMock.Verify(
            s => s.FindComboBox(
                It.Is<UiRequest>(r => r.Locator!.AutomationId == "cmbinbound"),
                null,
                null,
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void SelectComboBoxUia_WithoutSession_RoutesToNativeServiceWithNullContext()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);
        nativeMock
            .Setup(s => s.SelectComboBox(
                It.Is<UiRequest>(r => r.Operation == "selectcomboboxuia"),
                null,
                null,
                It.IsAny<CancellationToken>()))
            .Returns(new
            {
                operation = "selectcomboboxuia",
                success = false,
                stage = "no-search-context"
            });

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

        Assert.NotNull(result);
        nativeMock.Verify(
            s => s.SelectComboBox(
                It.Is<UiRequest>(r => r.Value == "Between"),
                null,
                null,
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public void FindComboBoxUia_InvalidLocator_ReturnsStructuredFailureFromService()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);
        nativeMock
            .Setup(s => s.FindComboBox(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new
            {
                operation = "findcomboboxuia",
                found = false,
                success = false,
                stage = "combo-not-found",
                error = "Native UIA resolver could not find a ComboBox for the locator."
            });

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var request = new UiRequest
        {
            Operation = "findcomboboxuia",
            Locator = new UiLocator { AutomationId = "missing-combo", ControlType = "ComboBox" }
        };

        var result = service.Execute(request);

        Assert.NotNull(result);
        nativeMock.Verify(
            s => s.FindComboBox(
                It.Is<UiRequest>(r => r.Locator!.AutomationId == "missing-combo"),
                null,
                null,
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Theory]
    [InlineData(null, 8000)]
    [InlineData(5000, 5000)]
    [InlineData(20000, 15000)]
    public void FindComboBoxUia_TimeoutMs_IsForwardedToService(int? requestTimeout, int expectedTimeout)
    {
        int? capturedTimeout = null;

        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);
        nativeMock
            .Setup(s => s.FindComboBox(
                It.IsAny<UiRequest>(),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Callback<UiRequest, IntPtr?, int?, CancellationToken>((req, _, _, _) => capturedTimeout = req.TimeoutMs)
            .Returns(new { operation = "findcomboboxuia", found = false, success = false, stage = "no-search-context" });

        var ctxMock = new Mock<IUiSessionContext>();
        ctxMock.Setup(c => c.ActiveSession).Returns((AutomationSession?)null);
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        service.Execute(new UiRequest
        {
            Operation = "findcomboboxuia",
            Locator = new UiLocator { AutomationId = "cmbinbound", ControlType = "ComboBox" },
            TimeoutMs = requestTimeout
        });

        Assert.Equal(requestTimeout, capturedTimeout);
        Assert.Equal(expectedTimeout, NativeUiaTimeoutPolicy.Resolve(requestTimeout));
    }
}
