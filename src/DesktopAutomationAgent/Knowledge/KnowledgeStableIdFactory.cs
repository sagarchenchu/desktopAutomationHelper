using System.Text.RegularExpressions;

namespace DesktopAutomationAgent.Knowledge;

/// <summary>
/// Deterministic semantic IDs. Excludes runtime refs, HWNDs, coordinates, and timestamps.
/// Element identity prefers AutomationId/ControlType/ClassName over Name-only keys.
/// </summary>
public sealed class KnowledgeStableIdFactory
{
    private static readonly Regex HwNdPattern = new(
        @"0x[0-9a-fA-F]+|hwnd:\d+|pid:\d+",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private readonly List<KnowledgeIdCollision> _collisions = [];
    private readonly Dictionary<string, string> _windowKeysById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _stateKeysById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _elementKeysById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _transitionKeysById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _workflowKeysById = new(StringComparer.Ordinal);

    public IReadOnlyList<KnowledgeIdCollision> Collisions => _collisions;

    public string WindowId(string? title, string? className, bool isModal)
    {
        var key = NormalizeKey(
            "window",
            Normalize(title),
            Normalize(className),
            isModal ? "modal" : "modeless");
        return Register("window", key, _windowKeysById, "win");
    }

    public string StateId(string semanticFingerprint)
    {
        var fp = Normalize(semanticFingerprint).ToLowerInvariant();
        if (fp.Length != 64 || fp.Any(c => c is < '0' or > '9' and (< 'a' or > 'f')))
            throw new ArgumentException("semanticFingerprint must be a 64-char lowercase hex SHA-256.", nameof(semanticFingerprint));

        var key = NormalizeKey("state", fp);
        return Register("state", key, _stateKeysById, "st");
    }

    public string ElementId(KnowledgeElementIdentity identity)
    {
        ArgumentNullException.ThrowIfNull(identity);

        var automationId = Normalize(identity.AutomationId);
        var controlType = Normalize(identity.ControlType);
        var className = Normalize(identity.ClassName);
        var name = Normalize(identity.Name);
        var role = Normalize(identity.Role);

        if (string.IsNullOrEmpty(controlType))
            throw new ArgumentException("Element controlType is required for stable IDs.", nameof(identity));

        // Prefer stronger attributes; Name alone is insufficient when AutomationId/ClassName exist.
        string key;
        if (!string.IsNullOrEmpty(automationId))
        {
            key = NormalizeKey("element", "aid", automationId, controlType, className);
        }
        else if (!string.IsNullOrEmpty(className))
        {
            key = NormalizeKey("element", "cls", className, controlType, name, role);
        }
        else
        {
            key = NormalizeKey("element", "name", controlType, name, role);
        }

        return Register("element", key, _elementKeysById, "el");
    }

    public string TransitionId(
        string sourceStateId,
        string normalizedAction,
        string? targetStateId,
        string outcome)
    {
        var key = NormalizeKey(
            "transition",
            Normalize(sourceStateId),
            StripRuntime(normalizedAction),
            Normalize(targetStateId),
            Normalize(outcome));
        return Register("transition", key, _transitionKeysById, "tr");
    }

    public string ConflictGroupId(string sourceStateId, string normalizedAction)
    {
        var hex = KnowledgeHash.Sha256Hex(NormalizeKey(
            "conflict",
            Normalize(sourceStateId),
            StripRuntime(normalizedAction)));
        return "cg_" + KnowledgeHash.StableIdPrefix(hex, 16);
    }

    public string WorkflowId(string entryStateId, IReadOnlyList<string> transitionIds)
    {
        var parts = new List<string> { "workflow", Normalize(entryStateId) };
        parts.AddRange(transitionIds.Select(Normalize));
        var key = NormalizeKey(parts.ToArray());
        return Register("workflow", key, _workflowKeysById, "wf");
    }

    public static string NormalizeActionKey(string actionKey)
    {
        if (string.IsNullOrWhiteSpace(actionKey))
            return string.Empty;

        return StripRuntime(actionKey.Trim());
    }

    public static string StripRuntime(string value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        var stripped = HwNdPattern.Replace(value, "{runtime}");
        // Drop ordinal/depth suffixes like |o3|d1 that are snapshot-local.
        stripped = Regex.Replace(
            stripped,
            @"\|o\d+\|d\d+",
            string.Empty,
            RegexOptions.CultureInvariant);
        return stripped.Trim();
    }

    private string Register(
        string kind,
        string key,
        Dictionary<string, string> keysById,
        string prefix)
    {
        var hex = KnowledgeHash.Sha256Hex(key);
        var id = $"{prefix}_{KnowledgeHash.StableIdPrefix(hex, 16)}";

        if (keysById.TryGetValue(id, out var existing) && !string.Equals(existing, key, StringComparison.Ordinal))
        {
            _collisions.Add(new KnowledgeIdCollision
            {
                Kind = kind,
                StableId = id,
                LeftKey = existing,
                RightKey = key
            });
        }
        else
        {
            keysById[id] = key;
        }

        return id;
    }

    private static string Normalize(string? value) =>
        string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

    private static string NormalizeKey(params string[] parts) =>
        string.Join('\u001f', parts);
}
