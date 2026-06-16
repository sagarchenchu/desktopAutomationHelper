using DesktopAutomationDriver.Models.Request;

namespace DesktopAutomationDriver.Services.NativeUia;

public interface INativeUiaComboBoxService
{
    object SelectComboBox(UiRequest request, IntPtr? activeWindowHwnd, int? processId);
}
