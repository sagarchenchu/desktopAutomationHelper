using DesktopAutomationDriver.Models.Recording;

namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>
/// Shared operation resolution for playback and Assistive BDD mapping.
/// Behavior must remain identical to the historical PlaybackService mapping.
/// </summary>
public static class RecordedActionOperationResolver
{
    public static string? ResolveOperation(RecordedAction action)
    {
        if (!string.IsNullOrWhiteSpace(action.Operation))
            return action.Operation;

        return action.ActionType switch
        {
            ActionType.Click => ResolveClickOperation(action),
            ActionType.MenuPathClick => "clicklogicalmenupath",
            ActionType.DoubleClick => "doubleclick",
            ActionType.RightClick => "rightclick",
            ActionType.Hover => "hover",
            ActionType.Select => "select",
            ActionType.Type => IsSendKeysValue(action.Value) ? "sendkeys" : "type",
            ActionType.TypeAndSelect => "typeandselect",
            ActionType.IsVisible => "isvisible",
            ActionType.IsClickable => "isclickable",
            ActionType.IsEnabled => "isenabled",
            ActionType.IsDisabled => null,
            ActionType.IsEditable => "iseditable",
            ActionType.GetTableHeaders => "gettableheaders",
            ActionType.GetTableData => "gettabledata",
            ActionType.IsChecked => "ischecked",
            ActionType.SelectCheckBox => ResolveCheckOperation(action),
            ActionType.ClearText => "clear",
            ActionType.GetValue => "getvalue",
            ActionType.Expand => "click",
            ActionType.Collapse => "click",
            ActionType.Maximize => "maximize",
            ActionType.Minimize => "minimize",
            ActionType.CloseWindow => "closewindow",
            ActionType.SwitchWindow => "switchwindow",
            ActionType.SetValue => "type",
            ActionType.Scroll => "scroll",
            _ => null
        };
    }

    public static string ResolveClickOperation(RecordedAction action)
    {
        if (string.Equals(action.Value, "{ESC}", StringComparison.OrdinalIgnoreCase) ||
            action.Description?.Contains("Cancel", StringComparison.OrdinalIgnoreCase) == true)
        {
            return "alertcancel";
        }

        if (action.Description?.Contains("Click OK", StringComparison.OrdinalIgnoreCase) == true)
            return "alertok";

        return IsSendKeysValue(action.Value) ? "sendkeys" : "click";
    }

    public static string? ResolveCheckOperation(RecordedAction action)
    {
        if (action.Description?.Contains("Uncheck", StringComparison.OrdinalIgnoreCase) == true)
            return "uncheck";

        return "check";
    }

    public static bool IsSendKeysValue(string? value) =>
        !string.IsNullOrWhiteSpace(value)
        && value.StartsWith('{')
        && value.EndsWith('}');
}
