using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Configuration;

/// <summary>
/// Canonical Jira-key contract shared by suite validation, agent defaults, and
/// schema drift checks. Matches the Assistive Mode / BDD action-map authority:
/// <c>^[A-Z][A-Z0-9_]{0,31}-[1-9][0-9]{0,15}$</c>.
/// </summary>
/// <remarks>
/// Suite files must already be uppercase and must not contain surrounding
/// whitespace (aligned with <c>suite.schema.json</c>). CLI <c>validate-keys</c>
/// may trim surrounding whitespace before validation. Interactive Assistive
/// input may trim and uppercase through the driver. This type does not
/// reference the driver.
/// </remarks>
public static class JiraKeyContract
{
    /// <summary>Authoritative Jira-key regex (Assistive Mode / bdd-action-map).</summary>
    public const string CanonicalPattern = @"^[A-Z][A-Z0-9_]{0,31}-[1-9][0-9]{0,15}$";

    /// <summary>Hard length bound (matches driver Assistive rules).</summary>
    public const int MaxLength = 64;

    /// <summary>Bounded match timeout for culture-invariant compiled matching.</summary>
    public static readonly TimeSpan MatchTimeout = TimeSpan.FromMilliseconds(250);

    private static readonly Regex CanonicalRegex = new(
        CanonicalPattern,
        RegexOptions.Compiled | RegexOptions.CultureInvariant,
        MatchTimeout);

    /// <summary>
    /// Compiles a project-specific additional restriction with the same options
    /// and timeout as the canonical matcher.
    /// </summary>
    public static Regex CompileProjectPattern(string pattern) =>
        new(pattern, RegexOptions.Compiled | RegexOptions.CultureInvariant, MatchTimeout);

    /// <summary>Returns true when <paramref name="key"/> matches the canonical contract.</summary>
    public static bool IsCanonical(string key)
    {
        if (string.IsNullOrEmpty(key) || key.Length > MaxLength)
            return false;

        try
        {
            return CanonicalRegex.IsMatch(key);
        }
        catch (RegexMatchTimeoutException)
        {
            return false;
        }
    }

    /// <summary>
    /// Validates a Jira key against the canonical contract, then an optional
    /// precompiled project-specific pattern.
    /// </summary>
    /// <param name="raw">Raw key from a suite file or CLI.</param>
    /// <param name="projectRegex">
    /// Optional additional restriction compiled from <c>Suites:JiraKeyPattern</c>.
    /// When null, only the canonical contract applies.
    /// </param>
    /// <param name="normalizedKey">Validated key on success (trimmed only when <paramref name="trimSurroundingWhitespace"/> is true).</param>
    /// <param name="error">Human-readable failure distinguishing canonical vs project rejection.</param>
    /// <param name="trimSurroundingWhitespace">
    /// When true (CLI <c>validate-keys</c>), surrounding whitespace is trimmed before matching.
    /// When false (suite files), surrounding whitespace is rejected so runtime matches JSON Schema.
    /// </param>
    public static bool TryValidate(
        string? raw,
        Regex? projectRegex,
        out string normalizedKey,
        out string error,
        bool trimSurroundingWhitespace = false)
    {
        normalizedKey = string.Empty;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(raw))
        {
            error = "jiraKey is required.";
            return false;
        }

        string key;
        if (trimSurroundingWhitespace)
        {
            key = raw.Trim();
        }
        else
        {
            // Suite files must match JSON Schema exactly — no silent trim.
            if (!string.Equals(raw, raw.Trim(), StringComparison.Ordinal))
            {
                error =
                    $"invalid jiraKey '{raw}': leading or trailing whitespace is not allowed " +
                    $"(canonical pattern: {CanonicalPattern}).";
                return false;
            }

            key = raw;
        }

        if (key.Length > MaxLength)
        {
            error =
                $"invalid jiraKey '{key}': does not match canonical Jira syntax " +
                $"({CanonicalPattern}); key exceeds {MaxLength} characters.";
            return false;
        }

        bool canonicalMatch;
        try
        {
            canonicalMatch = CanonicalRegex.IsMatch(key);
        }
        catch (RegexMatchTimeoutException)
        {
            canonicalMatch = false;
        }

        if (!canonicalMatch)
        {
            error =
                $"invalid jiraKey '{key}': does not match canonical Jira syntax ({CanonicalPattern}).";
            return false;
        }

        if (projectRegex is not null)
        {
            bool projectMatch;
            try
            {
                projectMatch = projectRegex.IsMatch(key);
            }
            catch (RegexMatchTimeoutException)
            {
                projectMatch = false;
            }

            if (!projectMatch)
            {
                error =
                    $"invalid jiraKey '{key}': rejected by configured project-specific " +
                    $"pattern ({projectRegex}).";
                return false;
            }
        }

        normalizedKey = key;
        return true;
    }
}
