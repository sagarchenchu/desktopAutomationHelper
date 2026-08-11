using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace DesktopAutomationDriver.Services.Assistive;

/// <summary>
/// Deterministic Phase-3-compatible pageId generation from window titles.
/// Must stay aligned with Agent ObjectRepository identifier constraints.
/// </summary>
public static class DeterministicPageIdGenerator
{
    public const int MaxLength = 64;
    public const string UntitledFallback = "untitled-window";

    private static readonly Regex IdentifierPattern = new(
        @"^[a-z][a-z0-9-]{0,63}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex NonKebabChars = new(@"[^a-z0-9-]+", RegexOptions.Compiled);
    private static readonly Regex CollapseHyphens = new(@"-{2,}", RegexOptions.Compiled);

    public static string FromWindowTitle(string? title, IReadOnlyDictionary<string, string>? titleToPageId = null)
    {
        var normalizedTitle = NormalizeTitle(title);
        if (titleToPageId != null
            && titleToPageId.TryGetValue(normalizedTitle, out var existing)
            && !string.IsNullOrWhiteSpace(existing))
        {
            return existing;
        }

        var used = titleToPageId?.Values.ToHashSet(StringComparer.Ordinal) ?? new HashSet<string>(StringComparer.Ordinal);
        return Allocate(normalizedTitle, title ?? string.Empty, used);
    }

    public static string NormalizeTitle(string? title)
    {
        if (string.IsNullOrWhiteSpace(title))
            return string.Empty;

        var normalized = title.Normalize(NormalizationForm.FormKC).Trim();
        normalized = Regex.Replace(normalized, @"\s+", " ");
        return normalized.ToLowerInvariant();
    }

    public static string Allocate(string normalizedTitle, string rawTitle, HashSet<string> usedPageIds)
    {
        var seed = string.IsNullOrWhiteSpace(normalizedTitle) ? UntitledFallback : normalizedTitle;
        var baseId = Sanitize(seed);
        var candidate = baseId;
        var suffix = 2;
        while (!usedPageIds.Add(candidate))
        {
            // Collision between different titles that sanitize identically.
            var hashSeed = rawTitle + "\u001f" + normalizedTitle + "\u001f" + suffix.ToString(CultureInfo.InvariantCulture);
            var suffixText = "-" + ShortHash(hashSeed);
            var maxBase = MaxLength - suffixText.Length;
            var truncated = Fit(baseId, Math.Max(1, maxBase), hashSeed);
            candidate = truncated + suffixText;
            if (candidate.Length > MaxLength)
                candidate = candidate[..MaxLength];
            suffix++;
        }

        return candidate;
    }

    private static string Sanitize(string seed)
    {
        var kebab = ToKebab(seed);
        if (string.IsNullOrWhiteSpace(kebab))
            kebab = UntitledFallback;

        if (char.IsDigit(kebab[0]))
            kebab = "p-" + kebab;

        if (!char.IsAsciiLetterLower(kebab[0]))
            kebab = "p-" + kebab.TrimStart('-');

        kebab = CollapseHyphens.Replace(kebab, "-").Trim('-');
        if (string.IsNullOrWhiteSpace(kebab) || !char.IsAsciiLetterLower(kebab[0]))
            kebab = UntitledFallback;

        if (kebab.Length > MaxLength)
            kebab = Fit(kebab, MaxLength, seed);

        if (!IdentifierPattern.IsMatch(kebab))
            kebab = Fit(UntitledFallback + "-" + ToKebab(seed), MaxLength, seed);

        return kebab;
    }

    private static string Fit(string value, int maxLength, string hashSeed)
    {
        if (maxLength < 1)
            return "p";

        if (value.Length <= maxLength && IdentifierPattern.IsMatch(value))
            return value;

        const int hashLength = 8;
        if (maxLength <= hashLength + 1)
        {
            var tiny = "p" + ShortHash(hashSeed)[..Math.Max(0, maxLength - 1)];
            return tiny[..Math.Min(tiny.Length, maxLength)];
        }

        var hash = ShortHash(hashSeed);
        var prefixLength = maxLength - hashLength - 1;
        var prefix = value.Length <= prefixLength ? value : value[..prefixLength];
        prefix = prefix.TrimEnd('-');
        if (string.IsNullOrWhiteSpace(prefix) || !char.IsAsciiLetterLower(prefix[0]))
            prefix = "p";

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
