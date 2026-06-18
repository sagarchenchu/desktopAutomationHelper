using DesktopAutomationDriver.Services;
using DesktopAutomationDriver.Services.NativeUia;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace DesktopAutomationDriver.Tests;

internal static class UiServiceTestFactory
{
    public static UiService Create(
        IUiSessionContext ctx,
        Mock<INativeUiaComboBoxService>? comboMock = null,
        Mock<INativeUiaBasicOperationService>? basicMock = null)
    {
        comboMock ??= new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);
        basicMock ??= new Mock<INativeUiaBasicOperationService>(MockBehavior.Strict);

        return new UiService(
            ctx,
            NullLogger<UiService>.Instance,
            comboMock.Object,
            basicMock.Object);
    }
}
