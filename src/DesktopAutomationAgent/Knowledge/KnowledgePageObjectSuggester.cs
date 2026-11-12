using System.Text.Json;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Workspace;

namespace DesktopAutomationAgent.Knowledge;

/// <summary>
/// Writes candidate Page Object suggestions only under object-repository/candidates.
/// Never writes active pages or modifies repository.json.
/// </summary>
public sealed class KnowledgePageObjectSuggester
{
    private readonly IWorkspaceManager _workspace;
    private readonly ObjectArtifactWriter _artifactWriter;

    public KnowledgePageObjectSuggester(IWorkspaceManager workspace, ObjectArtifactWriter artifactWriter)
    {
        _workspace = workspace;
        _artifactWriter = artifactWriter;
    }

    public IReadOnlyList<string> Suggest(ApplicationKnowledgeDocument knowledge)
    {
        ArgumentNullException.ThrowIfNull(knowledge);
        _workspace.EnsureInitialized();

        var candidatesRoot = Path.Combine(_workspace.RootPath, "object-repository", "candidates");
        Directory.CreateDirectory(candidatesRoot);
        KnowledgePathSafety.EnsureWritablePathInside(candidatesRoot, _workspace.RootPath);

        // Refuse to touch active pages or repository.json.
        var pagesRoot = Path.Combine(_workspace.RootPath, "object-repository", "pages");
        var repositoryJson = Path.Combine(_workspace.RootPath, "object-repository", "repository.json");
        _ = pagesRoot;
        _ = repositoryJson;

        var written = new List<string>();
        var byWindow = knowledge.Graph.Elements
            .Select(e => e)
            .GroupBy(_ => "page", StringComparer.Ordinal);

        // One candidate page per window that has associated elements via transitions.
        var elementsByWindow = new Dictionary<string, List<KnowledgeElementNode>>(StringComparer.Ordinal);
        foreach (var window in knowledge.Graph.Windows)
            elementsByWindow[window.WindowId] = [];

        var stateWindow = knowledge.Graph.States.ToDictionary(s => s.StateId, s => s.WindowId, StringComparer.Ordinal);
        var elements = knowledge.Graph.Elements.ToDictionary(e => e.ElementId, e => e, StringComparer.Ordinal);
        foreach (var transition in knowledge.Graph.Transitions)
        {
            if (transition.ElementId is null || !elements.TryGetValue(transition.ElementId, out var element))
                continue;
            if (!stateWindow.TryGetValue(transition.SourceStateId, out var windowId))
                continue;
            if (!elementsByWindow.TryGetValue(windowId, out var list))
            {
                list = [];
                elementsByWindow[windowId] = list;
            }

            if (list.All(e => e.ElementId != element.ElementId))
                list.Add(element);
        }

        foreach (var window in knowledge.Graph.Windows.OrderBy(w => w.WindowId, StringComparer.Ordinal))
        {
            if (!elementsByWindow.TryGetValue(window.WindowId, out var windowElements) || windowElements.Count == 0)
                continue;

            var pageId = SanitizePageId(window.Title, window.WindowId);
            var pageName = string.IsNullOrWhiteSpace(window.Title) ? pageId : window.Title;
            var document = new PageObjectDocument
            {
                SchemaVersion = ObjectRepositoryValidator.SchemaVersion,
                PageId = pageId,
                Name = pageName,
                State = "candidate",
                Elements = new Dictionary<string, ObjectElementDefinition>(StringComparer.Ordinal)
            };

            var used = new HashSet<string>(StringComparer.Ordinal);
            foreach (var element in windowElements.OrderBy(e => e.ElementId, StringComparer.Ordinal))
            {
                var elementId = ResolveElementId(element, used);
                var locator = BuildLocator(element);
                if (locator is null)
                    continue;

                document.Elements[elementId] = new ObjectElementDefinition
                {
                    Description = $"Suggested from knowledge {knowledge.KnowledgeId}",
                    Locator = locator,
                    Quality = new ObjectQuality
                    {
                        Grade = string.IsNullOrWhiteSpace(element.Identity.AutomationId) ? "weak" : "medium",
                        Warnings = ["knowledge-candidate-suggestion"]
                    },
                    Source = new ObjectSource
                    {
                        Kind = "capture",
                        Path = $"knowledge/{knowledge.ApplicationId}/{knowledge.BuildFingerprint}/{knowledge.KnowledgeId}"
                    }
                };
            }

            if (document.Elements.Count == 0)
                continue;

            var relative = Path.Combine(
                "object-repository",
                "candidates",
                knowledge.ApplicationId,
                knowledge.BuildFingerprint,
                $"{pageId}.page.json");
            var fullPath = Path.Combine(_workspace.RootPath, relative);
            KnowledgePathSafety.EnsureWritablePathInside(fullPath, Path.Combine(_workspace.RootPath, "object-repository", "candidates"));

            if (File.Exists(fullPath))
            {
                // Do not overwrite existing candidates; skip deterministically.
                continue;
            }

            var json = JsonSerializer.Serialize(document, KnowledgeJson.Indented);
            _artifactWriter.WriteAtomic(fullPath, json.EndsWith('\n') ? json : json + "\n");
            written.Add(relative.Replace('\\', '/'));
        }

        _ = byWindow;
        return written.OrderBy(p => p, StringComparer.Ordinal).ToList();
    }

    private static ObjectLocator? BuildLocator(KnowledgeElementNode element)
    {
        if (!string.IsNullOrWhiteSpace(element.Identity.AutomationId))
        {
            return new ObjectLocator
            {
                AutomationId = element.Identity.AutomationId,
                ControlType = string.IsNullOrWhiteSpace(element.Identity.ControlType)
                    ? null
                    : element.Identity.ControlType
            };
        }

        if (!string.IsNullOrWhiteSpace(element.Identity.Name)
            && !string.IsNullOrWhiteSpace(element.Identity.ControlType))
        {
            return new ObjectLocator
            {
                Name = element.Identity.Name,
                ControlType = element.Identity.ControlType
            };
        }

        return null;
    }

    private static string ResolveElementId(KnowledgeElementNode element, HashSet<string> used)
    {
        var raw = element.Identity.AutomationId
                  ?? element.Identity.Name
                  ?? element.ElementId;
        var kebab = ToKebab(raw);
        var candidate = kebab;
        var i = 2;
        while (!used.Add(candidate))
        {
            candidate = $"{kebab}-{i}";
            i++;
        }

        return candidate.Length <= 64 ? candidate : candidate[..64].Trim('-');
    }

    private static string SanitizePageId(string? title, string windowId)
    {
        var raw = string.IsNullOrWhiteSpace(title) ? windowId : title;
        var kebab = ToKebab(raw);
        if (string.IsNullOrWhiteSpace(kebab))
            kebab = "page";
        if (kebab.Length > 48)
            kebab = kebab[..48].Trim('-');
        return kebab;
    }

    private static string ToKebab(string value)
    {
        var chars = value.ToLowerInvariant().Select(ch =>
            char.IsLetterOrDigit(ch) ? ch : '-').ToArray();
        var collapsed = new string(chars);
        while (collapsed.Contains("--", StringComparison.Ordinal))
            collapsed = collapsed.Replace("--", "-", StringComparison.Ordinal);
        return collapsed.Trim('-');
    }
}
