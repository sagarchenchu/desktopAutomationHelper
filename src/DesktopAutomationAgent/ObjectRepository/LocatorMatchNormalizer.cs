using System.Net;
using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.ObjectRepository;

/// <summary>
/// Matching/normalization semantics aligned with the driver's NativeUiaText helpers
/// so capture uniqueness ranking matches finduia execution.
/// </summary>
public static class LocatorMatchNormalizer
{
    private static readonly Regex CollapseWhitespace = new(@"\s+", RegexOptions.Compiled);

    private static readonly Dictionary<string, int> ControlTypeIds = new(StringComparer.OrdinalIgnoreCase)
    {
        ["button"] = 50000,
        ["calendar"] = 50001,
        ["checkbox"] = 50002,
        ["combobox"] = 50003,
        ["edit"] = 50004,
        ["hyperlink"] = 50005,
        ["image"] = 50006,
        ["listitem"] = 50007,
        ["list"] = 50008,
        ["menu"] = 50011,
        ["menuitem"] = 50010,
        ["menubar"] = 50018,
        ["radiobutton"] = 50013,
        ["text"] = 50020,
        ["toolbar"] = 50021,
        ["tree"] = 50023,
        ["treeitem"] = 50024,
        ["custom"] = 50025,
        ["datagrid"] = 50028,
        ["dataitem"] = 50029,
        ["window"] = 50032,
        ["pane"] = 50033,
        ["tooltip"] = 50009
    };

    private static readonly Dictionary<int, string> ControlTypeNames = new()
    {
        [50000] = "Button",
        [50001] = "Calendar",
        [50002] = "CheckBox",
        [50003] = "ComboBox",
        [50004] = "Edit",
        [50005] = "Hyperlink",
        [50006] = "Image",
        [50007] = "ListItem",
        [50008] = "List",
        [50009] = "ToolTip",
        [50010] = "MenuItem",
        [50011] = "Menu",
        [50013] = "RadioButton",
        [50018] = "MenuBar",
        [50020] = "Text",
        [50021] = "ToolBar",
        [50023] = "Tree",
        [50024] = "TreeItem",
        [50025] = "Custom",
        [50028] = "DataGrid",
        [50029] = "DataItem",
        [50032] = "Window",
        [50033] = "Pane"
    };

    public static string Normalize(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        var decoded = WebUtility.HtmlDecode(text).Trim();
        decoded = decoded.Replace("&", string.Empty, StringComparison.Ordinal);
        decoded = CollapseWhitespace.Replace(decoded, " ");
        return decoded;
    }

    public static bool MatchesExact(string? candidate, string? requested) =>
        string.Equals(Normalize(candidate), Normalize(requested), StringComparison.OrdinalIgnoreCase);

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

        if (ControlTypeIds.TryGetValue(normalized, out var knownId))
            return knownId;

        if (int.TryParse(normalized, out var numericId))
            return numericId;

        var match = Regex.Match(normalized, @"\d+");
        if (match.Success && int.TryParse(match.Value, out numericId))
            return numericId;

        return null;
    }

    public static bool IsKnownControlType(string? controlType) =>
        ParseControlTypeId(controlType) is not null;

    public static string CanonicalControlType(string? controlType)
    {
        var id = ParseControlTypeId(controlType);
        if (id is null)
            return Normalize(controlType);

        return ControlTypeNames.TryGetValue(id.Value, out var name)
            ? name
            : $"ControlType({id.Value})";
    }

    public static string StrategyKey(string kind, params string?[] parts)
    {
        var normalized = parts.Select(Normalize).Select(static part => part.ToLowerInvariant());
        return kind + ":" + string.Join('|', normalized);
    }
}
