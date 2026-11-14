namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Fail-closed destructive detection for coordinate/vision actions and candidate labels.
/// </summary>
public static class DestructiveActionHeuristics
{
    private static readonly string[] DestructiveActions =
    [
        "delete", "close", "submit", "confirm", "destructive", "purchase", "pay", "buy",
        "checkout", "transfer", "remove", "uninstall", "format", "wipe", "sendmoney"
    ];

    private static readonly string[] DestructiveLabelTokens =
    [
        "delete", "remove", "uninstall", "wipe", "format",
        "pay", "payment", "purchase", "buy", "checkout", "order now",
        "submit", "confirm", "close", "transfer", "send money", "authorize payment"
    ];

    public static bool IsDestructiveAction(string? action)
    {
        if (string.IsNullOrWhiteSpace(action))
            return false;
        var normalized = action.Trim().ToLowerInvariant();
        return DestructiveActions.Any(token => normalized.Equals(token, StringComparison.Ordinal));
    }

    public static bool IsDestructiveLabel(string? label)
    {
        if (string.IsNullOrWhiteSpace(label))
            return false;

        var normalized = CollapseWhitespace(label).ToLowerInvariant();
        foreach (var token in DestructiveLabelTokens)
        {
            if (normalized.Equals(token, StringComparison.Ordinal)
                || normalized.Contains(token, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    public static bool IsDestructive(string? action, string? label, bool callerMarkedDestructive) =>
        callerMarkedDestructive
        || IsDestructiveAction(action)
        || IsDestructiveLabel(label);

    private static string CollapseWhitespace(string value) =>
        string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
