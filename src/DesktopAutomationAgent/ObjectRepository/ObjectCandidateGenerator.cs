using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectCandidateGenerator
{
    public const int MaxElementIdLength = 64;

    private static readonly Regex NonKebabChars = new(
        @"[^a-z0-9-]+",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex CollapseHyphens = new(
        @"-+",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex IdentifierPattern = new(
        @"^[a-z][a-z0-9-]{0,63}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public PageObjectDocument Generate(
        IReadOnlyList<JsonElement> nodes,
        string pageId,
        string pageName,
        string captureId)
    {
        ArgumentNullException.ThrowIfNull(nodes);

        var parsed = nodes
            .Select(ParseNode)
            .Where(static node => node is not null)
            .Select(static node => node!)
            .Where(HasUsableIdentity)
            .OrderBy(static node => node.SortKey, StringComparer.Ordinal)
            .ToList();

        var strategyCounts = BuildStrategyCounts(parsed);
        var elements = new Dictionary<string, ObjectElementDefinition>(StringComparer.Ordinal);
        var usedElementIds = new HashSet<string>(StringComparer.Ordinal);
        var unresolved = new List<Dictionary<string, object?>>();

        foreach (var node in parsed)
        {
            var locator = TryBuildLocator(node, strategyCounts);
            if (locator is null)
            {
                unresolved.Add(new Dictionary<string, object?>
                {
                    ["path"] = node.Path,
                    ["controlType"] = node.ControlType,
                    ["depth"] = node.Depth
                });
                continue;
            }

            var elementId = ResolveElementId(node, usedElementIds);
            elements[elementId] = new ObjectElementDefinition
            {
                Description = null,
                Locator = locator,
                Quality = new ObjectQuality
                {
                    Grade = GradeLocator(locator, parsed),
                    Warnings = []
                },
                Source = new ObjectSource
                {
                    Kind = "capture",
                    Path = $"captures/{pageId}/{captureId}.capture.json"
                }
            };
        }

        return new PageObjectDocument
        {
            SchemaVersion = ObjectRepositoryValidator.SchemaVersion,
            PageId = pageId,
            Name = pageName,
            State = "candidate",
            Elements = elements,
            Unresolved = unresolved.Count == 0
                ? null
                : JsonSerializer.SerializeToElement(unresolved)
        };
    }

    private static DumpNode? ParseNode(JsonElement element)
    {
        if (element.ValueKind != JsonValueKind.Object)
            return null;

        return new DumpNode(
            GetString(element, "path"),
            GetString(element, "automationId"),
            GetString(element, "name"),
            GetString(element, "controlType"),
            GetClassName(element),
            GetInt(element, "depth"));
    }

    private static string? GetString(JsonElement element, string propertyName) =>
        element.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private static int GetInt(JsonElement element, string propertyName) =>
        element.TryGetProperty(propertyName, out var value) && value.TryGetInt32(out var number)
            ? number
            : 0;

    private static string? GetClassName(JsonElement element) => GetString(element, "className");

    private static bool HasUsableIdentity(DumpNode node) =>
        HasValue(node.AutomationId)
        || (HasValue(node.Name) && HasValue(node.ControlType) && LocatorMatchNormalizer.IsKnownControlType(node.ControlType))
        || (HasValue(node.ClassName) && HasValue(node.ControlType) && LocatorMatchNormalizer.IsKnownControlType(node.ControlType));

    private static bool HasValue(string? value) => !string.IsNullOrWhiteSpace(value);

    private static Dictionary<string, int> BuildStrategyCounts(IReadOnlyList<DumpNode> nodes)
    {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var node in nodes)
        {
            foreach (var key in StrategyKeys(node))
            {
                counts.TryGetValue(key, out var count);
                counts[key] = count + 1;
            }
        }

        return counts;
    }

    private static IEnumerable<string> StrategyKeys(DumpNode node)
    {
        var automationId = LocatorMatchNormalizer.Normalize(node.AutomationId);
        var name = LocatorMatchNormalizer.Normalize(node.Name);
        var className = LocatorMatchNormalizer.Normalize(node.ClassName);
        var controlType = HasValue(node.ControlType) && LocatorMatchNormalizer.IsKnownControlType(node.ControlType)
            ? LocatorMatchNormalizer.CanonicalControlType(node.ControlType)
            : string.Empty;

        if (!string.IsNullOrEmpty(automationId) && !string.IsNullOrEmpty(controlType))
            yield return $"a:{automationId.ToLowerInvariant()}|ct:{controlType.ToLowerInvariant()}";

        if (!string.IsNullOrEmpty(automationId))
            yield return $"a:{automationId.ToLowerInvariant()}";

        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(controlType) && !string.IsNullOrEmpty(className))
        {
            yield return
                $"n:{name.ToLowerInvariant()}|ct:{controlType.ToLowerInvariant()}|cn:{className.ToLowerInvariant()}";
        }

        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(controlType))
            yield return $"n:{name.ToLowerInvariant()}|ct:{controlType.ToLowerInvariant()}";

        if (!string.IsNullOrEmpty(className) && !string.IsNullOrEmpty(controlType))
            yield return $"cn:{className.ToLowerInvariant()}|ct:{controlType.ToLowerInvariant()}";
    }

    private static ObjectLocator? TryBuildLocator(DumpNode node, IReadOnlyDictionary<string, int> counts)
    {
        var knownControlType = HasValue(node.ControlType)
                               && LocatorMatchNormalizer.IsKnownControlType(node.ControlType);
        var canonicalControlType = knownControlType
            ? LocatorMatchNormalizer.CanonicalControlType(node.ControlType)
            : null;

        if (HasValue(node.AutomationId) && knownControlType)
        {
            var key =
                $"a:{LocatorMatchNormalizer.Normalize(node.AutomationId).ToLowerInvariant()}|ct:{canonicalControlType!.ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator
                {
                    AutomationId = node.AutomationId,
                    ControlType = canonicalControlType
                };
            }
        }

        if (HasValue(node.AutomationId))
        {
            var key = $"a:{LocatorMatchNormalizer.Normalize(node.AutomationId).ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator { AutomationId = node.AutomationId };
            }
        }

        if (HasValue(node.Name) && knownControlType && HasValue(node.ClassName))
        {
            var key =
                $"n:{LocatorMatchNormalizer.Normalize(node.Name).ToLowerInvariant()}|ct:{canonicalControlType!.ToLowerInvariant()}|cn:{LocatorMatchNormalizer.Normalize(node.ClassName).ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator
                {
                    Name = node.Name,
                    ControlType = canonicalControlType,
                    ClassName = node.ClassName
                };
            }
        }

        if (HasValue(node.Name) && knownControlType)
        {
            var key =
                $"n:{LocatorMatchNormalizer.Normalize(node.Name).ToLowerInvariant()}|ct:{canonicalControlType!.ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator
                {
                    Name = node.Name,
                    ControlType = canonicalControlType
                };
            }
        }

        if (HasValue(node.ClassName) && knownControlType)
        {
            var key =
                $"cn:{LocatorMatchNormalizer.Normalize(node.ClassName).ToLowerInvariant()}|ct:{canonicalControlType!.ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator
                {
                    ClassName = node.ClassName,
                    ControlType = canonicalControlType
                };
            }
        }

        return null;
    }

    private static string GradeLocator(ObjectLocator locator, IReadOnlyList<DumpNode> allNodes)
    {
        if (HasValue(locator.AutomationId))
        {
            var matches = allNodes.Count(node =>
                LocatorMatchNormalizer.MatchesExact(node.AutomationId, locator.AutomationId));
            if (matches == 1)
                return "strong";
        }

        return "medium";
    }

    internal static string ResolveElementId(DumpNode node, HashSet<string> usedElementIds)
    {
        var seed = HasValue(node.AutomationId)
            ? node.AutomationId!
            : $"{node.Name}-{node.ControlType}";

        var baseId = SanitizeElementId(seed);
        var candidate = baseId;
        var suffix = 2;
        while (!usedElementIds.Add(candidate))
        {
            var suffixText = $"-{suffix}";
            var maxBaseLength = MaxElementIdLength - suffixText.Length;
            if (maxBaseLength < 1)
                throw new InvalidOperationException("Unable to allocate a valid element id within the 64-character limit.");

            var truncatedBase = FitIdentifier(baseId, maxBaseLength, seed + suffixText);
            candidate = truncatedBase + suffixText;
            suffix++;
        }

        if (!IdentifierPattern.IsMatch(candidate))
            throw new InvalidOperationException($"Generated element id '{candidate}' is invalid.");

        return candidate;
    }

    // Exposed for tests that construct DumpNode-like inputs via public API.
    internal static string SanitizeElementIdForTests(string seed) => SanitizeElementId(seed);

    private static string SanitizeElementId(string seed)
    {
        var kebab = ToKebab(seed);
        if (string.IsNullOrWhiteSpace(kebab))
            kebab = "element";

        if (!char.IsAsciiLetterLower(kebab[0]))
            kebab = "e-" + kebab.TrimStart('-');

        kebab = CollapseHyphens.Replace(kebab, "-").Trim('-');
        if (string.IsNullOrWhiteSpace(kebab) || !char.IsAsciiLetterLower(kebab[0]))
            kebab = "element";

        if (kebab.Length > MaxElementIdLength)
            kebab = FitIdentifier(kebab, MaxElementIdLength, seed);

        if (!IdentifierPattern.IsMatch(kebab))
            kebab = FitIdentifier("element-" + ToKebab(seed), MaxElementIdLength, seed);

        return kebab;
    }

    private static string FitIdentifier(string value, int maxLength, string hashSeed)
    {
        if (maxLength < 1)
            return "e";

        if (value.Length <= maxLength && IdentifierPattern.IsMatch(value) && value.Length <= MaxElementIdLength)
            return value.Length <= maxLength ? value : value[..maxLength];

        const int hashLength = 8;
        if (maxLength <= hashLength + 1)
        {
            var tiny = "e" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(hashSeed)))
                .ToLowerInvariant()[..Math.Max(0, maxLength - 1)];
            return tiny[..Math.Min(tiny.Length, maxLength)];
        }

        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(hashSeed)))
            .ToLowerInvariant()[..hashLength];
        var prefixLength = maxLength - hashLength - 1;
        var prefix = value.Length <= prefixLength ? value : value[..prefixLength];
        prefix = prefix.TrimEnd('-');
        if (string.IsNullOrWhiteSpace(prefix) || !char.IsAsciiLetterLower(prefix[0]))
            prefix = "e";

        var fitted = $"{prefix}-{hash}";
        if (fitted.Length > maxLength)
            fitted = fitted[..maxLength];

        return fitted;
    }

    private static string ToKebab(string value)
    {
        var lowered = value.Trim().ToLowerInvariant().Replace('_', '-').Replace(' ', '-');
        return CollapseHyphens.Replace(NonKebabChars.Replace(lowered, "-"), "-").Trim('-');
    }

    internal sealed record DumpNode(
        string? Path,
        string? AutomationId,
        string? Name,
        string? ControlType,
        string? ClassName,
        int Depth)
    {
        public string SortKey =>
            string.Join(
                '\u001f',
                Path ?? string.Empty,
                AutomationId ?? string.Empty,
                ControlType ?? string.Empty,
                ClassName ?? string.Empty,
                Name ?? string.Empty);
    }
}
