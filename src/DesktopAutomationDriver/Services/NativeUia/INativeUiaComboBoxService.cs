using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.NativeUia;

public interface INativeUiaComboBoxService
{
    object SelectComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);

    object FindComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);

    object InspectComboBox(
        UiRequest request,
        IntPtr? activeWindowHwnd,
        int? processId,
        CancellationToken cancellationToken = default);
}
