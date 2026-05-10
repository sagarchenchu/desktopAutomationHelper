using FlaUI.Core.Definitions;

namespace DesktopAutomationDriver.Services;

internal static class DynamicMenuHelpers
{
    public static bool IsDropdownContainerType(ControlType controlType) =>
        controlType == ControlType.Menu ||
        controlType == ControlType.ToolBar ||
        controlType == ControlType.Pane ||
        controlType == ControlType.Custom ||
        controlType == ControlType.Window;
}
