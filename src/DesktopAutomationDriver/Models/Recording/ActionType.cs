namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// The type of automation action that was recorded.
/// </summary>
public enum ActionType
{
    Click,
    DoubleClick,
    Hover,
    Select,
    Type,
    TypeAndSelect,
    IsVisible,
    IsClickable,
    IsEnabled,
    IsDisabled,
    IsEditable,
    GetTableHeaders,
    Assert,
    IsChecked,
    SelectCheckBox
}
