namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Fail-closed destructive detection for coordinate/vision actions and candidate labels.
/// Unknown action strings are denied (allow-list only).
/// </summary>
public static class DestructiveActionHeuristics
{
    private static readonly HashSet<string> AllowedActions = new(StringComparer.OrdinalIgnoreCase)
    {
        "click",
        "hover",
        "doubleclick",
        "double-click"
    };

    private static readonly string[] DestructiveLabelTokens =
    [
        "delete", "remove", "uninstall", "wipe", "format",
        "pay", "payment", "purchase", "buy", "checkout", "order now",
        "submit", "confirm", "close", "transfer", "send money", "authorize payment",
        "send", "approve", "publish", "install"
    ];

    /// <summary>
    /// True when the action is not on the allow-list (includes blank, unknown, and destructive verbs).
    /// </summary>
    public static bool IsUnknownOrDisallowedAction(string? action)
    {
        if (string.IsNullOrWhiteSpace(action))
            return true;
        return !AllowedActions.Contains(action.Trim());
    }

    public static bool IsDestructiveAction(string? action) =>
        IsUnknownOrDisallowedAction(action);

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
        || IsUnknownOrDisallowedAction(action)
        || IsDestructiveLabel(label);

    private static string CollapseWhitespace(string value) =>
        string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
