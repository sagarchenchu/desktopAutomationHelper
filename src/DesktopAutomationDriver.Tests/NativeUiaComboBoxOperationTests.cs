using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class NativeUiaComboBoxOperationTests
{
    [Fact]
    public void SelectComboBoxNativeUia_DelegatesToService()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);
        nativeMock
            .Setup(s => s.SelectComboBoxNativeUia(
                It.Is<UiRequest>(r => r.Operation == "selectcomboboxuia"),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>()))
            .Returns(new { operation = "selectcomboboxuia", success = true });

        var ctxMock = new Mock<IUiSessionContext>();
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var ex = Assert.Throws<InvalidOperationException>(() => service.Execute(new UiRequest
        {
            Operation = "selectcomboboxuia",
            Locator = new UiLocator { AutomationId = "cmbinbound", ControlType = "ComboBox" },
            Value = "Inbound"
        }));

        Assert.Contains("No active session", ex.Message);
    }

    [Theory]
    [InlineData("selectuia")]
    [InlineData("selectnativecombo")]
    [InlineData("selectcomboboxnative")]
    public void SelectComboBoxAliases_AreRegistered(string operation)
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>();
        nativeMock
            .Setup(s => s.SelectComboBoxNativeUia(It.IsAny<UiRequest>(), It.IsAny<IntPtr?>(), It.IsAny<int?>()))
            .Returns(new { success = false, stage = "combo-not-found" });

        var ctxMock = new Mock<IUiSessionContext>();
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        Assert.Throws<InvalidOperationException>(() => service.Execute(new UiRequest
        {
            Operation = operation,
            Locator = new UiLocator { AutomationId = "cmb" },
            Value = "X"
        }));
    }
}
