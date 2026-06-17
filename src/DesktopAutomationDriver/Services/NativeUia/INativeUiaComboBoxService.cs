using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.NativeUia;

public interface INativeUiaComboBoxService
{
    object SelectComboBoxNativeUia(UiRequest request, IntPtr? activeWindowHwnd, int? processId);

    object InspectComboBoxNativeUia(UiRequest request, IntPtr? activeWindowHwnd, int? processId);
}
