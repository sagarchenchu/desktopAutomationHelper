using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Driver;

/// <summary>
/// Removes bearer tokens and authorization headers from text before logging or displaying.
/// </summary>
public static class SecretRedactor
{
    private static readonly Regex BearerHeaderRegex = new(
        @"(?i)(Authorization\s*[:=]\s*)Bearer\s+\S+",
        RegexOptions.Compiled);

    private static readonly Regex BearerTokenRegex = new(
        @"(?i)\bBearer\s+[A-Za-z0-9+/=._-]{8,}",
        RegexOptions.Compiled);

    private static readonly Regex JsonTokenRegex = new(
        @"(?i)(""(?:token|bearerToken|authorizationHeader)""\s*:\s*)""[^""]*""",
        RegexOptions.Compiled);

    private static readonly Regex AssignmentTokenRegex = new(
        @"(?i)\b((?:token|bearerToken|authorizationHeader)\s*[:=]\s*)\S+",
        RegexOptions.Compiled);

    public static string Redact(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return text ?? string.Empty;

        var result = BearerHeaderRegex.Replace(text, "$1Bearer [REDACTED]");
        result = BearerTokenRegex.Replace(result, "Bearer [REDACTED]");
        result = JsonTokenRegex.Replace(result, "$1\"[REDACTED]\"");
        result = AssignmentTokenRegex.Replace(result, "$1[REDACTED]");
        return result;
    }

    public static string RedactKnownSecrets(string? text, params string?[] secrets)
    {
        var result = Redact(text);
        foreach (var secret in secrets)
        {
            if (string.IsNullOrWhiteSpace(secret) || secret.Length < 4)
                continue;

            result = result.Replace(secret, "[REDACTED]", StringComparison.Ordinal);
        }

        return result;
    }
}
