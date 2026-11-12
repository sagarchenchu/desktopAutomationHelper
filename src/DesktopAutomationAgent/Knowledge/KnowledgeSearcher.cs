namespace DesktopAutomationAgent.Knowledge;

/// <summary>
/// Deterministic lexical search only — no embeddings or vector DB.
/// </summary>
public sealed class KnowledgeSearcher
{
    public IReadOnlyList<KnowledgeSearchHit> Search(ApplicationKnowledgeDocument document, string query, int limit = 50)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentException.ThrowIfNullOrWhiteSpace(query);

        var terms = query
            .Split([' ', '\t', '\r', '\n', ',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(t => t.ToLowerInvariant())
            .Where(t => t.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(t => t, StringComparer.Ordinal)
            .ToArray();

        if (terms.Length == 0)
            return [];

        var hits = new List<KnowledgeSearchHit>();

        foreach (var window in document.Graph.Windows)
        {
            var text = $"{window.WindowId} {window.Title} {window.ClassName} {window.Role}";
            TryAdd(hits, "window", window.WindowId, text, terms);
        }

        foreach (var state in document.Graph.States)
        {
            var text = $"{state.StateId} {state.Label} {state.SemanticFingerprint} {state.WindowId}";
            TryAdd(hits, "state", state.StateId, text, terms);
        }

        foreach (var element in document.Graph.Elements)
        {
            var text =
                $"{element.ElementId} {element.Identity.AutomationId} {element.Identity.Name} {element.Identity.ControlType} {element.Identity.ClassName} {string.Join(' ', element.Capabilities)} {element.ApprovedObjectRef}";
            TryAdd(hits, "element", element.ElementId, text, terms);
        }

        foreach (var transition in document.Graph.Transitions)
        {
            var text =
                $"{transition.TransitionId} {transition.Action} {transition.ActionKind} {transition.ActionCategory} {transition.Outcome} {transition.SourceStateId} {transition.TargetStateId} {transition.ConflictGroupId}";
            TryAdd(hits, "transition", transition.TransitionId, text, terms);
        }

        foreach (var workflow in document.Workflows)
        {
            var text = $"{workflow.WorkflowId} {workflow.Name} {workflow.EntryStateId} {workflow.TerminalStateId}";
            TryAdd(hits, "workflow", workflow.WorkflowId, text, terms);
        }

        return hits
            .OrderByDescending(h => h.Score)
            .ThenBy(h => h.Kind, StringComparer.Ordinal)
            .ThenBy(h => h.Id, StringComparer.Ordinal)
            .Take(Math.Clamp(limit, 1, 500))
            .ToList();
    }

    private static void TryAdd(
        List<KnowledgeSearchHit> hits,
        string kind,
        string id,
        string text,
        string[] terms)
    {
        var haystack = text.ToLowerInvariant();
        var score = 0;
        foreach (var term in terms)
        {
            if (haystack.Contains(term, StringComparison.Ordinal))
                score += 1 + (haystack.StartsWith(term, StringComparison.Ordinal) ? 1 : 0);
        }

        if (score <= 0)
            return;

        hits.Add(new KnowledgeSearchHit
        {
            Kind = kind,
            Id = id,
            Text = text.Length <= 240 ? text : text[..240],
            Score = score
        });
    }
}
