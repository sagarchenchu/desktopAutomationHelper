namespace DesktopAutomationAgent.Exploration;

/// <summary>
/// Deterministic frontier of permitted navigation actions. Never re-queues an already
/// recorded state/action pair for the same crawl.
/// </summary>
public sealed class FrontierQueue
{
    private readonly SortedSet<FrontierEntry> _queue = new(FrontierComparer.Instance);
    private readonly HashSet<string> _seenPairs = new(StringComparer.Ordinal);
    private readonly HashSet<string> _queuedKeys = new(StringComparer.Ordinal);

    public int Count => _queue.Count;

    public int ExecutedPairCount => _seenPairs.Count;

    public bool HasExecuted(string stateFingerprint, string actionKey) =>
        _seenPairs.Contains(PairKey(stateFingerprint, actionKey));

    public bool TryEnqueue(ExplorationActionCandidate candidate)
    {
        ArgumentNullException.ThrowIfNull(candidate);
        if (string.IsNullOrWhiteSpace(candidate.ActionKey) || string.IsNullOrWhiteSpace(candidate.StateFingerprint))
            return false;

        var pair = PairKey(candidate.StateFingerprint, candidate.ActionKey);
        if (_seenPairs.Contains(pair) || !_queuedKeys.Add(pair))
            return false;

        _queue.Add(new FrontierEntry(candidate));
        return true;
    }

    public int EnqueueMany(IEnumerable<ExplorationActionCandidate> candidates)
    {
        var added = 0;
        foreach (var candidate in candidates)
        {
            if (TryEnqueue(candidate))
                added++;
        }

        return added;
    }

    public bool TryDequeue(out ExplorationActionCandidate? candidate)
    {
        candidate = null;
        if (_queue.Count == 0)
            return false;

        var next = _queue.Min!;
        _queue.Remove(next);
        var pair = PairKey(next.Candidate.StateFingerprint, next.Candidate.ActionKey);
        _queuedKeys.Remove(pair);
        _seenPairs.Add(pair);
        candidate = next.Candidate;
        return true;
    }

    public IReadOnlyList<ExplorationActionCandidate> Snapshot() =>
        _queue.Select(e => e.Candidate).ToList();

    public void MarkExecuted(string stateFingerprint, string actionKey) =>
        _seenPairs.Add(PairKey(stateFingerprint, actionKey));

    private static string PairKey(string stateFingerprint, string actionKey) =>
        stateFingerprint + "\u001f" + actionKey;

    private sealed record FrontierEntry(ExplorationActionCandidate Candidate);

    private sealed class FrontierComparer : IComparer<FrontierEntry>
    {
        public static FrontierComparer Instance { get; } = new();

        public int Compare(FrontierEntry? x, FrontierEntry? y)
        {
            if (ReferenceEquals(x, y))
                return 0;
            if (x is null)
                return -1;
            if (y is null)
                return 1;

            var cx = x.Candidate;
            var cy = y.Candidate;

            var byPriority = cx.Priority.CompareTo(cy.Priority);
            if (byPriority != 0)
                return byPriority;

            var byDepth = cx.Depth.CompareTo(cy.Depth);
            if (byDepth != 0)
                return byDepth;

            var byCategory = string.CompareOrdinal(cx.Category, cy.Category);
            if (byCategory != 0)
                return byCategory;

            var byElement = string.CompareOrdinal(cx.ElementKey, cy.ElementKey);
            if (byElement != 0)
                return byElement;

            var byKind = string.CompareOrdinal(cx.Kind.ToString(), cy.Kind.ToString());
            if (byKind != 0)
                return byKind;

            return string.CompareOrdinal(cx.ActionKey, cy.ActionKey);
        }
    }
}
