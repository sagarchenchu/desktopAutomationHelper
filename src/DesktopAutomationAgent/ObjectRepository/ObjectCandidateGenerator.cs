using System.Text.Json;
using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectCandidateGenerator
{
    private static readonly Regex NonKebabChars = new(
        @"[^a-z0-9-]+",
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
        || (HasValue(node.Name) && HasValue(node.ControlType))
        || (HasValue(node.ClassName) && HasValue(node.ControlType));

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
        if (HasValue(node.AutomationId) && HasValue(node.ControlType))
            yield return $"a:{node.AutomationId}|ct:{node.ControlType}";

        if (HasValue(node.AutomationId))
            yield return $"a:{node.AutomationId}";

        if (HasValue(node.Name) && HasValue(node.ControlType) && HasValue(node.ClassName))
            yield return $"n:{node.Name}|ct:{node.ControlType}|cn:{node.ClassName}";

        if (HasValue(node.Name) && HasValue(node.ControlType))
            yield return $"n:{node.Name}|ct:{node.ControlType}";

        if (HasValue(node.ClassName) && HasValue(node.ControlType))
            yield return $"cn:{node.ClassName}|ct:{node.ControlType}";
    }

    private static ObjectLocator? TryBuildLocator(DumpNode node, IReadOnlyDictionary<string, int> counts)
    {
        if (HasValue(node.AutomationId) && HasValue(node.ControlType))
        {
            var key = $"a:{node.AutomationId}|ct:{node.ControlType}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator
                {
                    AutomationId = node.AutomationId,
                    ControlType = node.ControlType
                };
            }
        }

        if (HasValue(node.AutomationId))
        {
            var key = $"a:{node.AutomationId}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator { AutomationId = node.AutomationId };
            }
        }

        if (HasValue(node.Name) && HasValue(node.ControlType) && HasValue(node.ClassName))
        {
            var key = $"n:{node.Name}|ct:{node.ControlType}|cn:{node.ClassName}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator
                {
                    Name = node.Name,
                    ControlType = node.ControlType,
                    ClassName = node.ClassName
                };
            }
        }

        if (HasValue(node.Name) && HasValue(node.ControlType))
        {
            var key = $"n:{node.Name}|ct:{node.ControlType}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator
                {
                    Name = node.Name,
                    ControlType = node.ControlType
                };
            }
        }

        if (HasValue(node.ClassName) && HasValue(node.ControlType))
        {
            var key = $"cn:{node.ClassName}|ct:{node.ControlType}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new ObjectLocator
                {
                    ClassName = node.ClassName,
                    ControlType = node.ControlType
                };
            }
        }

        return null;
    }

    private static string GradeLocator(ObjectLocator locator, IReadOnlyList<DumpNode> allNodes)
    {
        if (HasValue(locator.AutomationId))
        {
            var automationId = locator.AutomationId!;
            var matches = allNodes.Count(node =>
                string.Equals(node.AutomationId, automationId, StringComparison.Ordinal));
            if (matches == 1)
                return "strong";
        }

        return "medium";
    }

    private static string ResolveElementId(DumpNode node, HashSet<string> usedElementIds)
    {
        var baseId = HasValue(node.AutomationId)
            ? ToKebab(node.AutomationId!)
            : ToKebab($"{node.Name}-{node.ControlType}");

        if (string.IsNullOrWhiteSpace(baseId))
            baseId = "element";

        var candidate = baseId;
        var suffix = 2;
        while (!usedElementIds.Add(candidate))
        {
            candidate = $"{baseId}-{suffix}";
            suffix++;
        }

        return candidate;
    }

    private static string ToKebab(string value)
    {
        var lowered = value.Trim().ToLowerInvariant().Replace('_', '-').Replace(' ', '-');
        return NonKebabChars.Replace(lowered, "-").Trim('-');
    }

    private sealed record DumpNode(
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
