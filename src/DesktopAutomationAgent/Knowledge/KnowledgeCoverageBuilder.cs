namespace DesktopAutomationAgent.Knowledge;

public static class KnowledgeCoverageBuilder
{
    public static KnowledgeCoverageReport Build(
        ApplicationStateGraph graph,
        IEnumerable<string>? deniedSamples = null,
        IEnumerable<string>? truncationOrBudgetLimits = null)
    {
        ArgumentNullException.ThrowIfNull(graph);

        var successful = 0;
        var noOp = 0;
        var failed = 0;
        var conflicts = 0;

        foreach (var transition in graph.Transitions)
        {
            switch (NormalizeOutcome(transition.Outcome))
            {
                case "success":
                case "newwindow":
                case "closedwindow":
                case "modaldialog":
                    successful += transition.ObservationCount;
                    break;
                case "noop":
                    noOp += transition.ObservationCount;
                    break;
                case "failure":
                case "scopeviolation":
                case "stabilizationtimeout":
                case "staleref":
                    failed += transition.ObservationCount;
                    break;
            }

            if (!string.IsNullOrWhiteSpace(transition.ConflictGroupId))
                conflicts++;
        }

        return new KnowledgeCoverageReport
        {
            Windows = graph.Windows.Count,
            UniqueStates = graph.States.Count,
            ActionableElements = graph.Elements.Count,
            PermittedTransitionsAttempted = graph.Transitions.Sum(t => t.ObservationCount),
            SuccessfulTransitions = successful,
            NoOpTransitions = noOp,
            FailedTransitions = failed,
            DeniedSamples = deniedSamples?.Count() ?? 0,
            UnresolvedElements = graph.UnresolvedElements.Count + graph.AmbiguousElements.Count,
            Conflicts = conflicts,
            TruncationOrBudgetLimits = truncationOrBudgetLimits?
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(s => s, StringComparer.Ordinal)
                .ToList() ?? []
        };
    }

    private static string NormalizeOutcome(string? outcome) =>
        (outcome ?? string.Empty).Replace("_", string.Empty, StringComparison.Ordinal).ToLowerInvariant();
}
