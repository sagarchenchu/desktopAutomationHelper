using System.Text.Json;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Loads optional approved recovery manifests (equivalent transitions, approved candidates, fingerprints).
/// Missing files yield empty/no-op — tiers simply do not activate.
/// </summary>
public sealed class RecoveryManifestLoader
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly IWorkspaceManager _workspace;

    public RecoveryManifestLoader(IWorkspaceManager workspace)
    {
        _workspace = workspace;
    }

    public IReadOnlyDictionary<string, ApprovedEquivalentTransition> LoadEquivalentTransitions()
    {
        var path = Path.Combine(_workspace.RootPath, "knowledge", "approved", "recovery", "equivalent-transitions.json");
        if (!File.Exists(path))
            return new Dictionary<string, ApprovedEquivalentTransition>(StringComparer.Ordinal);

        try
        {
            var doc = JsonSerializer.Deserialize<EquivalentTransitionFile>(File.ReadAllText(path), JsonOptions);
            var map = new Dictionary<string, ApprovedEquivalentTransition>(StringComparer.Ordinal);
            foreach (var item in doc?.Transitions ?? [])
            {
                if (string.IsNullOrWhiteSpace(item.TransitionId) || string.IsNullOrWhiteSpace(item.FromAction))
                    continue;
                map[item.TransitionId] = item;
            }

            return map;
        }
        catch
        {
            return new Dictionary<string, ApprovedEquivalentTransition>(StringComparer.Ordinal);
        }
    }

    public (IReadOnlyDictionary<string, ObjectLocator> Locators, IReadOnlySet<string> Ids) LoadApprovedCandidates()
    {
        var root = Path.Combine(_workspace.RootPath, "object-repository", "candidates", "recovery");
        var locators = new Dictionary<string, ObjectLocator>(StringComparer.Ordinal);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        if (!Directory.Exists(root))
            return (locators, ids);

        foreach (var dir in Directory.EnumerateDirectories(root))
        {
            var approvedMarker = Path.Combine(dir, "approved.json");
            var repairPath = Path.Combine(dir, "candidate-repair.json");
            if (!File.Exists(approvedMarker) || !File.Exists(repairPath))
                continue;

            try
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(repairPath));
                var rootEl = doc.RootElement;
                if (!rootEl.TryGetProperty("objectReference", out var refEl)
                    || refEl.ValueKind != JsonValueKind.String)
                {
                    continue;
                }

                var objectRef = refEl.GetString();
                if (string.IsNullOrWhiteSpace(objectRef))
                    continue;

                if (!rootEl.TryGetProperty("newLocator", out var locatorEl)
                    || locatorEl.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                var locator = JsonSerializer.Deserialize<ObjectLocator>(locatorEl.GetRawText(), JsonOptions);
                if (locator is null)
                    continue;

                var validation = ObjectLocatorValidator.Validate(locator, objectRef + ".approvedCandidate");
                if (!validation.IsValid)
                    continue;

                locators[objectRef] = locator;
                ids.Add(objectRef);
            }
            catch
            {
                // skip malformed candidate
            }
        }

        return (locators, ids);
    }

    public string? LoadExpectedApplicationFingerprint()
    {
        var path = Path.Combine(_workspace.RootPath, "knowledge", "approved", "recovery", "application-fingerprint.txt");
        if (!File.Exists(path))
            return null;

        try
        {
            var text = File.ReadAllText(path).Trim();
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }
        catch
        {
            return null;
        }
    }

    private sealed class EquivalentTransitionFile
    {
        public List<ApprovedEquivalentTransition>? Transitions { get; set; }
    }
}
