namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>Pure BDD statement validation for Assistive recording.</summary>
public static class BddStatementRules
{
    public const int MaxLength = 2000;

    public static bool TryNormalize(string? raw, out string statement, out string error)
    {
        statement = string.Empty;
        error = string.Empty;

        if (raw is null)
        {
            error = "BDD statement is required.";
            return false;
        }

        var trimmed = raw.Trim();
        if (trimmed.Length == 0)
        {
            error = "BDD statement cannot be empty.";
            return false;
        }

        if (trimmed.Length > MaxLength)
        {
            error = $"BDD statement must be at most {MaxLength} characters.";
            return false;
        }

        // Treat as one logical line: reject embedded CR/LF rather than silently rewriting content.
        if (trimmed.Contains('\r') || trimmed.Contains('\n'))
        {
            error = "BDD statement must be a single line.";
            return false;
        }

        if (trimmed.Any(char.IsControl))
        {
            error = "BDD statement contains invalid control characters.";
            return false;
        }

        statement = trimmed;
        return true;
    }
}
