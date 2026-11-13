using DesktopAutomationDriver.Services.Assistive;

namespace DesktopAutomationDriver.Services.Debug;

/// <summary>
/// Generates candidate locators using existing object-repository / Phase 3 identity rules.
/// Never promotes candidates.
/// </summary>
public sealed class LocatorCandidateGenerator
{
    public IReadOnlyList<LocatorCandidate> Generate(
        DebugUiaNode target,
        IReadOnlyList<DebugUiaNode>? siblingsForUniqueness = null)
    {
        ArgumentNullException.ThrowIfNull(target);

        var pool = new List<DebugUiaNode> { target };
        if (siblingsForUniqueness is { Count: > 0 })
            pool.AddRange(siblingsForUniqueness);

        var counts = BuildStrategyCounts(pool);
        var locator = TryBuildLocator(target, counts);
        if (locator is null)
            return [];

        if (!Phase3LocatorContract.IsValid(locator, out var error))
        {
            return
            [
                new LocatorCandidate
                {
                    ElementId = DeterministicElementIdGenerator.Sanitize(
                        !string.IsNullOrWhiteSpace(target.AutomationId)
                            ? target.AutomationId!
                            : $"{target.Name}-{target.ControlType}"),
                    Locator = locator,
                    Grade = "invalid",
                    Strategy = "rejected",
                    Warnings = [error],
                    AutoPromoted = false
                }
            ];
        }

        var used = new HashSet<string>(StringComparer.Ordinal);
        var elementId = DeterministicElementIdGenerator.Resolve(
            new Models.Recording.ElementInfo
            {
                AutomationId = target.AutomationId,
                Name = target.Name,
                ControlType = target.ControlType,
                ClassName = target.ClassName
            },
            used);

        var strategy = DescribeStrategy(locator);
        var grade = Grade(locator, pool);
        var warnings = new List<string>();
        if (!locator.ContainsKey("automationId"))
            warnings.Add("Locator may be ambiguous without automationId; manual review required.");

        return
        [
            new LocatorCandidate
            {
                ElementId = elementId,
                Locator = locator,
                Grade = grade,
                Strategy = strategy,
                Warnings = warnings,
                AutoPromoted = false
            }
        ];
    }

    private static Dictionary<string, int> BuildStrategyCounts(IReadOnlyList<DebugUiaNode> nodes)
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

    private static IEnumerable<string> StrategyKeys(DebugUiaNode node)
    {
        var automationId = Normalize(node.AutomationId);
        var name = Normalize(node.Name);
        var className = Normalize(node.ClassName);
        var controlType = Phase3KnownControlTypes.IsKnown(node.ControlType)
            ? node.ControlType!.Trim()
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

    private static Dictionary<string, object?>? TryBuildLocator(
        DebugUiaNode node,
        IReadOnlyDictionary<string, int> counts)
    {
        var knownControlType = Phase3KnownControlTypes.IsKnown(node.ControlType);
        var canonicalControlType = knownControlType ? node.ControlType!.Trim() : null;
        var hasAutomationId = HasValue(node.AutomationId);
        var hasName = HasValue(node.Name);
        var hasClassName = HasValue(node.ClassName);

        if (hasAutomationId && knownControlType)
        {
            var key =
                $"a:{Normalize(node.AutomationId).ToLowerInvariant()}|ct:{canonicalControlType!.ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new Dictionary<string, object?>(StringComparer.Ordinal)
                {
                    ["automationId"] = node.AutomationId,
                    ["controlType"] = canonicalControlType
                };
            }
        }

        if (hasAutomationId)
        {
            var key = $"a:{Normalize(node.AutomationId).ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new Dictionary<string, object?>(StringComparer.Ordinal)
                {
                    ["automationId"] = node.AutomationId
                };
            }
        }

        if (hasName && knownControlType && hasClassName)
        {
            var key =
                $"n:{Normalize(node.Name).ToLowerInvariant()}|ct:{canonicalControlType!.ToLowerInvariant()}|cn:{Normalize(node.ClassName).ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new Dictionary<string, object?>(StringComparer.Ordinal)
                {
                    ["name"] = node.Name,
                    ["controlType"] = canonicalControlType,
                    ["className"] = node.ClassName
                };
            }
        }

        if (hasName && knownControlType)
        {
            var key =
                $"n:{Normalize(node.Name).ToLowerInvariant()}|ct:{canonicalControlType!.ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new Dictionary<string, object?>(StringComparer.Ordinal)
                {
                    ["name"] = node.Name,
                    ["controlType"] = canonicalControlType
                };
            }
        }

        if (hasClassName && knownControlType)
        {
            var key =
                $"cn:{Normalize(node.ClassName).ToLowerInvariant()}|ct:{canonicalControlType!.ToLowerInvariant()}";
            if (counts.TryGetValue(key, out var count) && count == 1)
            {
                return new Dictionary<string, object?>(StringComparer.Ordinal)
                {
                    ["className"] = node.ClassName,
                    ["controlType"] = canonicalControlType
                };
            }
        }

        return null;
    }

    private static string Grade(IReadOnlyDictionary<string, object?> locator, IReadOnlyList<DebugUiaNode> pool)
    {
        if (locator.TryGetValue("automationId", out var idObj) && idObj is string id && HasValue(id))
        {
            var matches = pool.Count(n =>
                string.Equals(Normalize(n.AutomationId), Normalize(id), StringComparison.OrdinalIgnoreCase));
            if (matches == 1)
                return "strong";
        }

        return "medium";
    }

    private static string DescribeStrategy(IReadOnlyDictionary<string, object?> locator)
    {
        var hasA = locator.ContainsKey("automationId");
        var hasN = locator.ContainsKey("name");
        var hasC = locator.ContainsKey("className");
        var hasT = locator.ContainsKey("controlType");
        if (hasA && hasT) return "automationId+controlType";
        if (hasA) return "automationId";
        if (hasN && hasT && hasC) return "name+controlType+className";
        if (hasN && hasT) return "name+controlType";
        if (hasC && hasT) return "className+controlType";
        return "unknown";
    }

    private static string Normalize(string? value) => (value ?? string.Empty).Trim();

    private static bool HasValue(string? value) => !string.IsNullOrWhiteSpace(value);
}
