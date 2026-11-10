namespace DesktopAutomationAgent.Observation;

internal static class PatternCapabilityMapper
{
    public static IReadOnlyList<string> Map(IEnumerable<string>? patterns)
    {
        var capabilities = new SortedSet<string>(StringComparer.Ordinal);
        if (patterns is null)
            return capabilities.ToList();

        foreach (var pattern in patterns)
        {
            switch (Normalize(pattern))
            {
                case "invoke":
                    capabilities.Add("invoke");
                    break;
                case "expandcollapse":
                    capabilities.Add("expandCollapse");
                    break;
                case "selectionitem":
                    capabilities.Add("selectItem");
                    break;
                case "selection":
                    capabilities.Add("selectionContainer");
                    break;
                case "toggle":
                    capabilities.Add("toggle");
                    break;
                case "value":
                    capabilities.Add("setValue");
                    capabilities.Add("readValue");
                    break;
                case "scroll":
                    capabilities.Add("scroll");
                    break;
                case "grid":
                    capabilities.Add("grid");
                    break;
                case "window":
                    capabilities.Add("window");
                    break;
            }
        }

        return capabilities.ToList();
    }

    public static bool IsActionable(IEnumerable<string> capabilities) =>
        capabilities.Any(static c => c is
            "invoke" or "expandCollapse" or "selectItem" or "selectionContainer"
            or "toggle" or "setValue" or "scroll" or "grid" or "window");

    private static string Normalize(string pattern) =>
        pattern.Trim().Replace(" ", string.Empty, StringComparison.Ordinal).ToLowerInvariant();
}
