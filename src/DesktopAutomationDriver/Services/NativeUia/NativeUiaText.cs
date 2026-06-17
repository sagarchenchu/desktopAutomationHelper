using System.Net;
using System.Text.RegularExpressions;

namespace DesktopAutomationDriver.Services.NativeUia;

/// <summary>
/// Text normalization and matching helpers for native UIA ComboBox item selection.
/// </summary>
internal static class NativeUiaText
{
    private static readonly Regex CollapseWhitespace = new(@"\s+", RegexOptions.Compiled);

    public static string Normalize(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        var decoded = WebUtility.HtmlDecode(text).Trim();
        decoded = decoded.Replace("&", string.Empty, StringComparison.Ordinal);
        decoded = CollapseWhitespace.Replace(decoded, " ");
        return decoded;
    }

    public static bool Matches(string? candidate, string? requested, string matchMode)
    {
        var left = Normalize(candidate);
        var right = Normalize(requested);

        if (string.IsNullOrEmpty(right))
            return string.IsNullOrEmpty(left);

        return matchMode.ToLowerInvariant() switch
        {
            "contains" => left.Contains(right, StringComparison.OrdinalIgnoreCase),
            "startswith" => left.StartsWith(right, StringComparison.OrdinalIgnoreCase),
            "regex" => Regex.IsMatch(left, requested!, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
            "exact" => string.Equals(left, right, StringComparison.OrdinalIgnoreCase),
            _ => string.Equals(left, right, StringComparison.OrdinalIgnoreCase)
        };
    }

    public static bool ValuesEquivalent(string? actual, string? requested)
    {
        if (Matches(actual, requested, "exact"))
            return true;

        var left = Normalize(actual);
        var right = Normalize(requested);
        return !string.IsNullOrEmpty(right)
               && left.Contains(right, StringComparison.OrdinalIgnoreCase);
    }

    public static string ControlTypeName(int controlTypeId) =>
        controlTypeId switch
        {
            50000 => "Button",
            50002 => "CheckBox",
            50003 => "ComboBox",
            50004 => "Edit",
            50007 => "ListItem",
            50008 => "List",
            50009 => "Menu",
            50010 => "MenuItem",
            50013 => "RadioButton",
            50020 => "Text",
            50024 => "Tree",
            50025 => "TreeItem",
            50028 => "DataGrid",
            50029 => "DataItem",
            50032 => "Window",
            50033 => "Pane",
            _ => $"ControlType({controlTypeId})"
        };

    public static int? ParseControlTypeId(string? controlType)
    {
        if (string.IsNullOrWhiteSpace(controlType))
            return null;

        return controlType.Trim().ToLowerInvariant() switch
        {
            "button" => 50000,
            "checkbox" => 50002,
            "combobox" => 50003,
            "edit" => 50004,
            "listitem" => 50007,
            "list" => 50008,
            "menu" => 50009,
            "menuitem" => 50010,
            "radiobutton" => 50013,
            "text" => 50020,
            "tree" => 50024,
            "treeitem" => 50025,
            "datagrid" => 50028,
            "dataitem" => 50029,
            "window" => 50032,
            "pane" => 50033,
            _ => int.TryParse(controlType, out var id) ? id : null
        };
    }
}
