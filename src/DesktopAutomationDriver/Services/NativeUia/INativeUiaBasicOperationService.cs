using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.NativeUia;

public interface INativeUiaBasicOperationService
{
    object Click(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);

    object Type(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);

    object SendKeys(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);

    object Clear(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);

    object Focus(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);
}
