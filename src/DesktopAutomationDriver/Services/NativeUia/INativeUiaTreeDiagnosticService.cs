using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.NativeUia;

public interface INativeUiaTreeDiagnosticService
{
    object DumpTree(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);

    object FindElement(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);
}
