using DesktopAutomationAgent.Observation;

namespace DesktopAutomationAgent.Exploration;

/// <summary>
/// Derives technical navigation candidates from a semantic snapshot and policy.
/// Deny-by-default: InvokePattern alone never authorizes a click.
/// </summary>
public sealed class CapabilityResolver
{
    private static readonly Dictionary<ExplorationActionKind, int> CategoryPriority = new()
    {
        [ExplorationActionKind.Expand] = 10,
        [ExplorationActionKind.Collapse] = 11,
        [ExplorationActionKind.TabSelect] = 20,
        [ExplorationActionKind.MenuOpen] = 30,
        [ExplorationActionKind.SelectItem] = 40,
        [ExplorationActionKind.Scroll] = 50,
        [ExplorationActionKind.Focus] = 60,
        [ExplorationActionKind.Hover] = 70,
        [ExplorationActionKind.Invoke] = 80
    };

    private readonly List<string> _deniedSamples = [];

    public IReadOnlyList<string> DeniedActionSamples => _deniedSamples;

    public IReadOnlyList<ExplorationActionCandidate> Resolve(
        SemanticSnapshot snapshot,
        string stateFingerprint,
        ExplorationPolicyDocument policy,
        int depth)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(policy);

        _deniedSamples.Clear();
        if (!string.Equals(policy.Mode, "navigate", StringComparison.OrdinalIgnoreCase))
            return [];

        var allowedCategories = new HashSet<string>(policy.AllowedActionCategories, StringComparer.Ordinal);
        var candidates = new List<ExplorationActionCandidate>();

        foreach (var node in snapshot.Nodes)
        {
            if (!node.Enabled || node.Offscreen)
                continue;

            foreach (var proposal in ProposeForNode(node))
            {
                if (!IsAuthorized(node, proposal.Kind, proposal.Category, policy, allowedCategories, out var denyReason))
                {
                    if (_deniedSamples.Count < 32 && denyReason is not null)
                        _deniedSamples.Add($"{proposal.Category}:{BuildElementKey(node)}:{denyReason}");
                    continue;
                }

                candidates.Add(new ExplorationActionCandidate
                {
                    ActionKey = BuildActionKey(stateFingerprint, proposal.Kind, node),
                    Kind = proposal.Kind,
                    Category = proposal.Category,
                    DriverOperation = proposal.DriverOperation,
                    ElementKey = BuildElementKey(node),
                    AutomationId = node.AutomationId,
                    Name = node.Name,
                    ControlType = node.ControlType,
                    ClassName = node.ClassName,
                    SnapshotId = snapshot.SnapshotId,
                    NodeRef = node.Ref,
                    StateFingerprint = stateFingerprint,
                    Depth = depth,
                    Priority = CategoryPriority[proposal.Kind],
                    Ordinal = node.Ordinal,
                    DriverArguments = proposal.Arguments
                });
            }
        }

        return candidates
            .OrderBy(c => c.Priority)
            .ThenBy(c => c.Category, StringComparer.Ordinal)
            .ThenBy(c => c.ElementKey, StringComparer.Ordinal)
            .ThenBy(c => c.Kind.ToString(), StringComparer.Ordinal)
            .ThenBy(c => c.ActionKey, StringComparer.Ordinal)
            .ToList();
    }

    private static IEnumerable<(ExplorationActionKind Kind, string Category, string DriverOperation, Dictionary<string, string>? Arguments)> ProposeForNode(
        SemanticNode node)
    {
        var capabilities = new HashSet<string>(node.Capabilities ?? [], StringComparer.Ordinal);
        var controlType = NormalizeControlType(node.ControlType);

        // expandtreeitem/collapsetreeitem are tree operations; menus use menuOpen instead.
        if (capabilities.Contains("expandCollapse")
            && controlType is "TreeItem" or "Tree")
        {
            yield return (ExplorationActionKind.Expand, "expandCollapse", "expandtreeitem", null);
            yield return (ExplorationActionKind.Collapse, "expandCollapse", "collapsetreeitem", null);
        }

        if (capabilities.Contains("selectItem")
            || controlType is "TabItem" or "ListItem")
        {
            if (controlType == "TabItem")
            {
                yield return (ExplorationActionKind.TabSelect, "tabSelect", "selecttabuia",
                    BuildValueArgs(node));
            }
            else
            {
                yield return (ExplorationActionKind.SelectItem, "selectItem", "select",
                    BuildValueArgs(node));
            }
        }

        // Only expandable submenu containers are MenuOpen. Leaf MenuItem commands
        // (Exit/Reset/Print/Archive/etc.) go through Invoke + explicit allow — never
        // via a global menuOpen category grant.
        if (IsExpandableMenuContainer(node, controlType, capabilities))
        {
            yield return (ExplorationActionKind.MenuOpen, "menuOpen", "clickuia", null);
        }

        if (capabilities.Contains("scroll"))
            yield return (ExplorationActionKind.Scroll, "scroll", "scroll",
                new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["direction"] = "down",
                    ["amount"] = "1",
                    ["mode"] = "pattern"
                });

        if (node.Focusable == true)
            yield return (ExplorationActionKind.Focus, "focus", "focusuia", null);

        yield return (ExplorationActionKind.Hover, "hover", "hover", null);

        // Invoke is proposed only as a technical candidate; authorization is separate.
        // Leaf menu commands surface here (not as MenuOpen).
        if (capabilities.Contains("invoke")
            || (controlType == "MenuItem" && !IsExpandableMenuContainer(node, controlType, capabilities)))
        {
            yield return (ExplorationActionKind.Invoke, "invoke", "clickuia", null);
        }
    }

    private static bool IsExpandableMenuContainer(
        SemanticNode node,
        string controlType,
        HashSet<string> capabilities)
    {
        if (controlType is "Menu" or "MenuBar")
            return true;

        if (controlType != "MenuItem" && !LooksLikeMenu(node))
            return false;

        if (capabilities.Contains("expandCollapse"))
            return true;

        if (node.Interaction?.ExpandCollapse is string expandState
            && !string.IsNullOrWhiteSpace(expandState))
            return true;

        // Top-level menu containers often named *Menu without ExpandCollapse projected yet.
        if (controlType == "MenuItem"
            && !string.IsNullOrWhiteSpace(node.AutomationId)
            && node.AutomationId.EndsWith("Menu", StringComparison.OrdinalIgnoreCase)
            && !node.AutomationId.EndsWith("MenuItem", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private static Dictionary<string, string>? BuildValueArgs(SemanticNode node)
    {
        if (!string.IsNullOrWhiteSpace(node.Name))
        {
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["value"] = node.Name!
            };
        }

        return null;
    }

    private static bool LooksLikeMenu(SemanticNode node) =>
        (!string.IsNullOrWhiteSpace(node.AutomationId)
         && node.AutomationId.Contains("Menu", StringComparison.OrdinalIgnoreCase))
        || (!string.IsNullOrWhiteSpace(node.Name)
            && node.Name.Contains("Menu", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Re-authorize a live node for an already-queued candidate. Fail closed when
    /// the live element no longer matches the intended action.
    /// </summary>
    public bool IsStillAuthorized(
        SemanticNode liveNode,
        ExplorationActionCandidate action,
        ExplorationPolicyDocument policy,
        out string? denyReason)
    {
        ArgumentNullException.ThrowIfNull(liveNode);
        ArgumentNullException.ThrowIfNull(action);
        ArgumentNullException.ThrowIfNull(policy);

        var allowedCategories = new HashSet<string>(policy.AllowedActionCategories, StringComparer.Ordinal);
        return IsAuthorized(liveNode, action.Kind, action.Category, policy, allowedCategories, out denyReason);
    }

    /// <summary>
    /// Stable identity key without ordinal/depth — used for unambiguous rematch only.
    /// </summary>
    public static string BuildStableIdentityKey(SemanticNode node)
    {
        var aid = string.IsNullOrWhiteSpace(node.AutomationId) ? "-" : node.AutomationId.Trim();
        var name = string.IsNullOrWhiteSpace(node.Name) ? "-" : node.Name.Trim();
        var ct = string.IsNullOrWhiteSpace(node.ControlType) ? "-" : NormalizeControlType(node.ControlType);
        var cls = string.IsNullOrWhiteSpace(node.ClassName) ? "-" : node.ClassName.Trim();
        return $"{ct}|{aid}|{name}|{cls}";
    }

    private bool IsAuthorized(
        SemanticNode node,
        ExplorationActionKind kind,
        string category,
        ExplorationPolicyDocument policy,
        HashSet<string> allowedCategories,
        out string? denyReason)
    {
        denyReason = null;

        if (ForbiddenActionVocabulary.Matches(node.Name)
            || ForbiddenActionVocabulary.Matches(node.AutomationId)
            || ForbiddenActionVocabulary.Matches(node.HelpText))
        {
            denyReason = "forbidden-vocabulary";
            return false;
        }

        if (MatchesAnyRule(node, kind, category, policy.DenyRules))
        {
            denyReason = "deny-rule";
            return false;
        }

        if (kind == ExplorationActionKind.Hover && !policy.AllowHover)
        {
            denyReason = "hover-disabled";
            return false;
        }

        if (!allowedCategories.Contains(category)
            && !MatchesAnyRule(node, kind, category, policy.AllowRules))
        {
            denyReason = "category-not-allowed";
            return false;
        }

        // Invoke/click requires an explicit allow rule even when the category is listed.
        if (kind == ExplorationActionKind.Invoke
            && !MatchesAnyRule(node, kind, category, policy.AllowRules))
        {
            denyReason = "invoke-requires-explicit-allow";
            return false;
        }

        if (!allowedCategories.Contains(category)
            && MatchesAnyRule(node, kind, category, policy.AllowRules))
        {
            // Explicit allow rule can authorize a specific action even when the category
            // list is narrow, except hover still needs allowHover.
            return true;
        }

        if (allowedCategories.Contains(category)
            && kind != ExplorationActionKind.Invoke)
        {
            return true;
        }

        if (kind == ExplorationActionKind.Invoke
            && allowedCategories.Contains(category)
            && MatchesAnyRule(node, kind, category, policy.AllowRules))
        {
            return true;
        }

        denyReason = "deny-by-default";
        return false;
    }

    private static bool MatchesAnyRule(
        SemanticNode node,
        ExplorationActionKind kind,
        string category,
        IEnumerable<ExplorationElementRule> rules)
    {
        foreach (var rule in rules)
        {
            if (!MatchesRule(node, kind, category, rule))
                continue;
            return true;
        }

        return false;
    }

    private static bool MatchesRule(
        SemanticNode node,
        ExplorationActionKind kind,
        string category,
        ExplorationElementRule rule)
    {
        if (!string.IsNullOrWhiteSpace(rule.AutomationId)
            && !string.Equals(rule.AutomationId, node.AutomationId, StringComparison.Ordinal))
            return false;

        if (!string.IsNullOrWhiteSpace(rule.Name)
            && !string.Equals(rule.Name, node.Name, StringComparison.Ordinal))
            return false;

        if (!string.IsNullOrWhiteSpace(rule.NameContains)
            && (node.Name is null
                || node.Name.IndexOf(rule.NameContains, StringComparison.OrdinalIgnoreCase) < 0))
            return false;

        if (!string.IsNullOrWhiteSpace(rule.ControlType)
            && !string.Equals(
                NormalizeControlType(rule.ControlType),
                NormalizeControlType(node.ControlType),
                StringComparison.OrdinalIgnoreCase))
            return false;

        if (!string.IsNullOrWhiteSpace(rule.ClassName)
            && !string.Equals(rule.ClassName, node.ClassName, StringComparison.Ordinal))
            return false;

        if (!string.IsNullOrWhiteSpace(rule.Action) && rule.Action != "*")
        {
            var expected = MapActionToken(rule.Action);
            if (expected is null || expected != kind)
                return false;
        }

        if (!string.IsNullOrWhiteSpace(rule.ActionCategory) && rule.ActionCategory != "*")
        {
            if (!string.Equals(rule.ActionCategory, category, StringComparison.Ordinal))
                return false;
        }

        return true;
    }

    private static ExplorationActionKind? MapActionToken(string action) =>
        action.Trim().ToLowerInvariant() switch
        {
            "expand" => ExplorationActionKind.Expand,
            "collapse" => ExplorationActionKind.Collapse,
            "tabselect" => ExplorationActionKind.TabSelect,
            "menuopen" => ExplorationActionKind.MenuOpen,
            "scroll" => ExplorationActionKind.Scroll,
            "focus" => ExplorationActionKind.Focus,
            "hover" => ExplorationActionKind.Hover,
            "selectitem" => ExplorationActionKind.SelectItem,
            "invoke" => ExplorationActionKind.Invoke,
            _ => null
        };

    public static string BuildElementKey(SemanticNode node)
    {
        var aid = string.IsNullOrWhiteSpace(node.AutomationId) ? "-" : node.AutomationId.Trim();
        var name = string.IsNullOrWhiteSpace(node.Name) ? "-" : node.Name.Trim();
        var ct = string.IsNullOrWhiteSpace(node.ControlType) ? "-" : NormalizeControlType(node.ControlType);
        return $"{ct}|{aid}|{name}|o{node.Ordinal}|d{node.Depth}";
    }

    private static string BuildActionKey(string stateFingerprint, ExplorationActionKind kind, SemanticNode node) =>
        $"{kind}:{BuildElementKey(node)}";

    internal static string NormalizeControlType(string? controlType)
    {
        if (string.IsNullOrWhiteSpace(controlType))
            return string.Empty;

        var trimmed = controlType.Trim();
        return trimmed switch
        {
            "ControlType(50019)" or "50019" => "TabItem",
            "ControlType(50024)" or "50024" => "TreeItem",
            "ControlType(50007)" or "50007" => "ListItem",
            "ControlType(50011)" or "50011" => "MenuItem",
            "ControlType(50009)" or "50009" => "Menu",
            "ControlType(50010)" or "50010" => "MenuBar",
            "ControlType(50033)" or "50033" => "Tree",
            "ControlType(50008)" or "50008" => "List",
            "ControlType(50032)" or "50032" => "Button",
            "ControlType(50004)" or "50004" => "Edit",
            "ControlType(50020)" or "50020" => "Tab",
            "ControlType(50037)" or "50037" => "TitleBar",
            "ControlType(50021)" or "50021" => "TitleBar",
            _ => trimmed
        };
    }
}
