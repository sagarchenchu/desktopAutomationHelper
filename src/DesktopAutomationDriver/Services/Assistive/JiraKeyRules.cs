using System.Text.RegularExpressions;

namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>Pure Jira key validation and canonicalization for Assistive recording.</summary>
public static class JiraKeyRules
{
    public const int MaxLength = 64;

    private static readonly Regex Pattern = new(
        @"^[A-Z][A-Z0-9_]{0,31}-[1-9][0-9]{0,15}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly char[] InvalidPathChars =
        ['/', '\\', '.', ' ', '\t', '\r', '\n', '"', '\'', '<', '>', '|', '?', '*', ':'];

    public static bool TryCanonicalize(string? raw, out string canonical, out string error)
    {
        canonical = string.Empty;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(raw))
        {
            error = "Jira key is required.";
            return false;
        }

        var trimmed = raw.Trim();
        if (trimmed.Length > MaxLength)
        {
            error = $"Jira key must be at most {MaxLength} characters.";
            return false;
        }

        if (trimmed.IndexOfAny(InvalidPathChars) >= 0
            || trimmed.Any(char.IsControl))
        {
            error = "Jira key contains invalid characters.";
            return false;
        }

        canonical = trimmed.ToUpperInvariant();
        if (!Pattern.IsMatch(canonical))
        {
            error = "Jira key must look like PROJECT-1234 (letters/digits/underscore, hyphen, positive number).";
            canonical = string.Empty;
            return false;
        }

        return true;
    }
}
