using System.Text.RegularExpressions;

namespace DesktopAutomationDriver.NativeUia;

internal static class NativeUiaText
{
    public static string Normalize(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        return Regex.Replace(
            value.Replace("\r", " ").Replace("\n", " "),
            @"\s+",
            " ").Trim();
    }

    public static bool TextMatches(string actual, string expected, string? matchMode)
    {
        actual = Normalize(actual);
        expected = Normalize(expected);

        if (string.IsNullOrWhiteSpace(expected))
            return false;

        return matchMode?.ToLowerInvariant() switch
        {
            "contains" => actual.Contains(expected, StringComparison.OrdinalIgnoreCase),
            "startswith" => actual.StartsWith(expected, StringComparison.OrdinalIgnoreCase),
            "endswith" => actual.EndsWith(expected, StringComparison.OrdinalIgnoreCase),
            "regex" => Regex.IsMatch(actual, expected, RegexOptions.IgnoreCase),
            _ => string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase)
        };
    }

    public static string ControlTypeName(int controlTypeId) => controlTypeId switch
    {
        NativeUiaConstants.UIA_ButtonControlTypeId => "Button",
        NativeUiaConstants.UIA_CheckBoxControlTypeId => "CheckBox",
        NativeUiaConstants.UIA_ComboBoxControlTypeId => "ComboBox",
        NativeUiaConstants.UIA_EditControlTypeId => "Edit",
        NativeUiaConstants.UIA_ListItemControlTypeId => "ListItem",
        NativeUiaConstants.UIA_ListControlTypeId => "List",
        NativeUiaConstants.UIA_MenuItemControlTypeId => "MenuItem",
        NativeUiaConstants.UIA_TextControlTypeId => "Text",
        NativeUiaConstants.UIA_WindowControlTypeId => "Window",
        NativeUiaConstants.UIA_PaneControlTypeId => "Pane",
        NativeUiaConstants.UIA_CustomControlTypeId => "Custom",
        NativeUiaConstants.UIA_DataItemControlTypeId => "DataItem",
        NativeUiaConstants.UIA_RadioButtonControlTypeId => "RadioButton",
        _ => $"ControlType({controlTypeId})"
    };
}
