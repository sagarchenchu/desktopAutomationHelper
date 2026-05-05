namespace DesktopAutomationDriver.Models.Recording;

/// <summary>
/// The type of automation action that was recorded.
/// </summary>
public enum ActionType
{
    Click,
    MenuPathClick,
    DoubleClick,
    RightClick,
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
    GetTableData,
    Assert,
    IsChecked,
    SelectCheckBox,
    ClearText,
    GetValue,
    Expand,
    Collapse,
    Maximize,
    Minimize,
    CloseWindow,
    SwitchWindow,
    SetValue,
    Scroll,
    DragAndDrop,
}
