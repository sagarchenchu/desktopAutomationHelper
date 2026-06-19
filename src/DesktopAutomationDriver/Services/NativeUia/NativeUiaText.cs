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
            50001 => "Calendar",
            50002 => "CheckBox",
            50003 => "ComboBox",
            50004 => "Edit",
            50005 => "Hyperlink",
            50006 => "Image",
            50007 => "ListItem",
            50008 => "List",
            50009 => "ToolTip",
            50010 => "MenuItem",
            50011 => "Menu",
            50013 => "RadioButton",
            50018 => "MenuBar",
            50020 => "Text",
            50021 => "ToolBar",
            50023 => "Tree",
            50024 => "TreeItem",
            50025 => "Custom",
            50026 => "Custom",
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

        var normalized = controlType
            .Trim()
            .Replace("_", "", StringComparison.Ordinal)
            .Replace("-", "", StringComparison.Ordinal)
            .Replace(" ", "", StringComparison.Ordinal)
            .ToLowerInvariant();

        return normalized switch
        {
            "button" => 50000,
            "calendar" => 50001,
            "checkbox" => 50002,
            "combobox" => 50003,
            "edit" => 50004,
            "hyperlink" => 50005,
            "image" => 50006,
            "listitem" => 50007,
            "list" => 50008,
            "menu" => 50011,
            "menuitem" => 50010,
            "menubar" => 50018,
            "radiobutton" => 50013,
            "text" => 50020,
            "toolbar" => 50021,
            "tree" => 50023,
            "treeitem" => 50024,
            "custom" => 50025,
            "datagrid" => 50028,
            "dataitem" => 50029,
            "window" => 50032,
            "pane" => 50033,
            "controltype(50011)" => 50011,
            "controltype50011" => 50011,
            "50011" => 50011,
            _ => TryParseControlTypeNumber(normalized)
        };
    }

    private static int? TryParseControlTypeNumber(string value)
    {
        if (int.TryParse(value, out var id))
            return id;

        var match = Regex.Match(value, @"\d+");
        if (match.Success && int.TryParse(match.Value, out id))
            return id;

        return null;
    }
}
