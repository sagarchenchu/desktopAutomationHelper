namespace DesktopAutomationAgent.Exploration;

/// <summary>
/// Additional deny-by-default guard. Labels matching these terms are always denied
/// unless a future higher-level policy mode explicitly authorizes them (not in Phase 4B).
/// String heuristics are never the sole authorization rule.
/// </summary>
public static class ForbiddenActionVocabulary
{
    private static readonly string[] Terms =
    [
        "delete",
        "remove",
        "approve",
        "submit",
        "save",
        "pay",
        "purchase",
        "send",
        "transfer",
        "publish",
        "install",
        "uninstall"
    ];

    public static IReadOnlyList<string> All => Terms;

    public static bool Matches(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var normalized = Normalize(text);
        foreach (var term in Terms)
        {
            if (normalized.Contains(term, StringComparison.Ordinal))
                return true;
        }

        return false;
    }

    public static string Normalize(string text)
    {
        var chars = text.Trim().ToLowerInvariant().ToCharArray();
        for (var i = 0; i < chars.Length; i++)
        {
            if (!char.IsLetterOrDigit(chars[i]))
                chars[i] = ' ';
        }

        return string.Join(' ', new string(chars).Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }
}
