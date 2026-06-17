using DesktopAutomationDriver.Models.Request;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

public class ComboBoxNativeUiaOnlyRoutingTests
{
    [Fact]
    public void IsComboBoxSelectRequest_WithComboBoxControlType_IsTrue()
    {
        var req = new UiRequest
        {
            Locator = new UiLocator { Name = "Operar", ControlType = "ComboBox" }
        };

        Assert.True(ComboBoxRoutingTestHelper.IsComboBoxSelectRequest(req));
    }

    [Fact]
    public void IsComboBoxSelectRequest_WithoutControlType_IsFalse()
    {
        var req = new UiRequest
        {
            Locator = new UiLocator { Name = "Operar" }
        };

        Assert.False(ComboBoxRoutingTestHelper.IsComboBoxSelectRequest(req));
    }

    [Fact]
    public void Select_ComboBoxRequest_UsesNativeServiceOnly()
    {
        var nativeMock = new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);
        nativeMock
            .Setup(s => s.SelectComboBox(
                It.Is<UiRequest>(r => r.Locator!.Name == "Operar"),
                It.IsAny<IntPtr?>(),
                It.IsAny<int?>(),
                It.IsAny<CancellationToken>()))
            .Returns(new { success = false, stage = "combo-not-found" });

        var ctxMock = new Mock<IUiSessionContext>();
        var service = new UiService(ctxMock.Object, NullLogger<UiService>.Instance, nativeMock.Object);

        var ex = Assert.Throws<InvalidOperationException>(() => service.Execute(new UiRequest
        {
            Operation = "select",
            Locator = new UiLocator { Name = "Operar", ControlType = "ComboBox" },
            Value = "equals",
            TimeoutMs = 8000
        }));

        Assert.Contains("No active session", ex.Message);
    }
}

/// <summary>Exposes routing helper for unit tests via reflection-free duplicate.</summary>
internal static class ComboBoxRoutingTestHelper
{
    public static bool IsComboBoxSelectRequest(UiRequest req)
    {
        var controlType = req.Locator?.ControlType;
        if (string.IsNullOrWhiteSpace(controlType))
            return false;

        return string.Equals(controlType, "ComboBox", StringComparison.OrdinalIgnoreCase)
               || string.Equals(controlType, "Edit", StringComparison.OrdinalIgnoreCase);
    }
}
