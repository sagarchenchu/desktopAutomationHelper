namespace DesktopAutomationAgent.Knowledge;

public sealed class KnowledgeComparer
{
    private readonly ApplicationFingerprintBuilder _fingerprintBuilder = new();

    public KnowledgeCompareResult Compare(ApplicationKnowledgeDocument left, ApplicationKnowledgeDocument right)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);

        var result = new KnowledgeCompareResult
        {
            LeftBuildFingerprint = left.BuildFingerprint,
            RightBuildFingerprint = right.BuildFingerprint
        };

        result.CompatibleFingerprints = _fingerprintBuilder.AreCompatible(
            left.Fingerprint,
            right.Fingerprint,
            out var mismatches);

        if (!result.CompatibleFingerprints)
        {
            result.Differences.Add(
                "incompatible-fingerprints: " + string.Join(", ", mismatches));
            result.GraphsEqual = false;
            return result;
        }

        var leftHash = KnowledgeArtifactWriter.ComputeContentHash(left);
        var rightHash = KnowledgeArtifactWriter.ComputeContentHash(right);
        result.GraphsEqual = string.Equals(leftHash, rightHash, StringComparison.Ordinal);

        if (!result.GraphsEqual)
        {
            DiffCounts("windows", left.Graph.Windows.Count, right.Graph.Windows.Count, result.Differences);
            DiffCounts("states", left.Graph.States.Count, right.Graph.States.Count, result.Differences);
            DiffCounts("elements", left.Graph.Elements.Count, right.Graph.Elements.Count, result.Differences);
            DiffCounts("transitions", left.Graph.Transitions.Count, right.Graph.Transitions.Count, result.Differences);
            DiffCounts("conflicts", left.Coverage.Conflicts, right.Coverage.Conflicts, result.Differences);

            var leftTransitions = left.Graph.Transitions.Select(t => t.TransitionId).ToHashSet(StringComparer.Ordinal);
            var rightTransitions = right.Graph.Transitions.Select(t => t.TransitionId).ToHashSet(StringComparer.Ordinal);
            foreach (var id in leftTransitions.Except(rightTransitions).OrderBy(x => x, StringComparer.Ordinal).Take(20))
                result.Differences.Add($"transition-only-left:{id}");
            foreach (var id in rightTransitions.Except(leftTransitions).OrderBy(x => x, StringComparer.Ordinal).Take(20))
                result.Differences.Add($"transition-only-right:{id}");

            if (result.Differences.Count == 0)
                result.Differences.Add("content-hash-mismatch");
        }

        return result;
    }

    private static void DiffCounts(string label, int left, int right, List<string> differences)
    {
        if (left != right)
            differences.Add($"{label}:{left}!={right}");
    }
}
