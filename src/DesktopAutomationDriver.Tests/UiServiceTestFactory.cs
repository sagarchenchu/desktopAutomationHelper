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
        Mock<INativeUiaBasicOperationService>? basicMock = null,
        Mock<INativeUiaTreeDiagnosticService>? treeMock = null)
    {
        comboMock ??= new Mock<INativeUiaComboBoxService>(MockBehavior.Strict);
        basicMock ??= new Mock<INativeUiaBasicOperationService>(MockBehavior.Strict);
        treeMock ??= new Mock<INativeUiaTreeDiagnosticService>(MockBehavior.Strict);

        return new UiService(
            ctx,
            NullLogger<UiService>.Instance,
            comboMock.Object,
            basicMock.Object,
            treeMock.Object);
    }
}
