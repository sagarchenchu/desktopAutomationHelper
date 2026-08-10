using System.Text.RegularExpressions;

namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>
/// Phase 3 ObjectRepository-recognized UIA control type names.
/// Must stay aligned with Agent <c>LocatorMatchNormalizer.IsKnownControlType</c>.
/// </summary>
public static class Phase3KnownControlTypes
{
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

    public static bool IsKnown(string? controlType) =>
        ParseControlTypeId(controlType) is not null;

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
}
