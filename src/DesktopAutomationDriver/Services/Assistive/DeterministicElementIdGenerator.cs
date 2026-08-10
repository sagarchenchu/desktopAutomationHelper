using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DesktopAutomationDriver.Models.Recording;

namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>
/// Deterministic Phase-3-compatible element id generation for Assistive page candidates.
/// Mirrors Agent ObjectCandidateGenerator sanitization rules.
/// </summary>
public static class DeterministicElementIdGenerator
{
    public const int MaxLength = 64;

    private static readonly Regex IdentifierPattern = new(
        @"^[a-z][a-z0-9-]{0,63}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex NonKebabChars = new(@"[^a-z0-9-]+", RegexOptions.Compiled);
    private static readonly Regex CollapseHyphens = new(@"-{2,}", RegexOptions.Compiled);

    public static string Resolve(ElementInfo? element, HashSet<string> usedElementIds)
    {
        var automationId = element?.AutomationId;
        var seed = !string.IsNullOrWhiteSpace(automationId)
            ? automationId!
            : $"{element?.Name}-{element?.ControlType}";

        var baseId = Sanitize(seed);
        var candidate = baseId;
        var suffix = 2;
        while (!usedElementIds.Add(candidate))
        {
            var suffixText = $"-{suffix}";
            var maxBaseLength = MaxLength - suffixText.Length;
            if (maxBaseLength < 1)
                throw new InvalidOperationException("Unable to allocate a valid element id within the 64-character limit.");

            var truncatedBase = Fit(baseId, maxBaseLength, seed + suffixText);
            candidate = truncatedBase + suffixText;
            suffix++;
        }

        if (!IdentifierPattern.IsMatch(candidate))
            throw new InvalidOperationException($"Generated element id '{candidate}' is invalid.");

        return candidate;
    }

    public static string Sanitize(string seed)
    {
        var kebab = ToKebab(seed);
        if (string.IsNullOrWhiteSpace(kebab))
            kebab = "element";

        if (!char.IsAsciiLetterLower(kebab[0]))
            kebab = "e-" + kebab.TrimStart('-');

        kebab = CollapseHyphens.Replace(kebab, "-").Trim('-');
        if (string.IsNullOrWhiteSpace(kebab) || !char.IsAsciiLetterLower(kebab[0]))
            kebab = "element";

        if (kebab.Length > MaxLength)
            kebab = Fit(kebab, MaxLength, seed);

        if (!IdentifierPattern.IsMatch(kebab))
            kebab = Fit("element-" + ToKebab(seed), MaxLength, seed);

        return kebab;
    }

    private static string Fit(string value, int maxLength, string hashSeed)
    {
        if (maxLength < 1)
            return "e";

        if (value.Length <= maxLength && IdentifierPattern.IsMatch(value))
            return value;

        const int hashLength = 8;
        if (maxLength <= hashLength + 1)
        {
            var tiny = "e" + ShortHash(hashSeed)[..Math.Max(0, maxLength - 1)];
            return tiny[..Math.Min(tiny.Length, maxLength)];
        }

        var hash = ShortHash(hashSeed);
        var prefixLength = maxLength - hashLength - 1;
        var prefix = value.Length <= prefixLength ? value : value[..prefixLength];
        prefix = prefix.TrimEnd('-');
        if (string.IsNullOrWhiteSpace(prefix) || !char.IsAsciiLetterLower(prefix[0]))
            prefix = "e";

        var fitted = $"{prefix}-{hash}";
        return fitted.Length <= maxLength ? fitted : fitted[..maxLength];
    }

    private static string ToKebab(string value)
    {
        var lowered = value.Trim().ToLowerInvariant().Replace('_', '-').Replace(' ', '-');
        return CollapseHyphens.Replace(NonKebabChars.Replace(lowered, "-"), "-").Trim('-');
    }

    private static string ShortHash(string seed) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(seed))).ToLowerInvariant()[..8];
}
