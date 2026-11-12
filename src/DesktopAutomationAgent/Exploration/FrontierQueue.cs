namespace DesktopAutomationAgent.Exploration;

/// <summary>
/// Deterministic frontier of permitted navigation actions. Never re-queues an already
/// recorded state/action/window pair for the same crawl.
/// </summary>
public sealed class FrontierQueue
{
    private readonly SortedSet<FrontierEntry> _queue = new(FrontierComparer.Instance);
    private readonly HashSet<string> _seenPairs = new(StringComparer.Ordinal);
    private readonly HashSet<string> _queuedKeys = new(StringComparer.Ordinal);

    public int Count => _queue.Count;

    public int ExecutedPairCount => _seenPairs.Count;

    public bool HasExecuted(ExplorationActionCandidate candidate)
    {
        ArgumentNullException.ThrowIfNull(candidate);
        return _seenPairs.Contains(PairKey(candidate));
    }

    public bool HasExecuted(
        string stateFingerprint,
        string actionKey,
        long? targetHwnd = null,
        int? targetProcessId = null) =>
        _seenPairs.Contains(PairKey(stateFingerprint, actionKey, targetHwnd, targetProcessId));

    public bool TryEnqueue(ExplorationActionCandidate candidate)
    {
        ArgumentNullException.ThrowIfNull(candidate);
        if (string.IsNullOrWhiteSpace(candidate.ActionKey) || string.IsNullOrWhiteSpace(candidate.StateFingerprint))
            return false;

        var pair = PairKey(candidate);
        if (_seenPairs.Contains(pair) || !_queuedKeys.Add(pair))
            return false;

        if (!_queue.Add(new FrontierEntry(candidate)))
        {
            // Pair key said unique but comparer treated entries as equal — fail closed.
            _queuedKeys.Remove(pair);
            throw new InvalidOperationException(
                $"Frontier SortedSet.Add failed for pair '{pair}' despite unique pair key; fail-closed.");
        }

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
        var pair = PairKey(next.Candidate);
        _queuedKeys.Remove(pair);
        _seenPairs.Add(pair);
        candidate = next.Candidate;
        return true;
    }

    public IReadOnlyList<ExplorationActionCandidate> Snapshot() =>
        _queue.Select(e => e.Candidate).ToList();

    public void MarkExecuted(ExplorationActionCandidate candidate)
    {
        ArgumentNullException.ThrowIfNull(candidate);
        _seenPairs.Add(PairKey(candidate));
    }

    public void MarkExecuted(
        string stateFingerprint,
        string actionKey,
        long? targetHwnd = null,
        int? targetProcessId = null) =>
        _seenPairs.Add(PairKey(stateFingerprint, actionKey, targetHwnd, targetProcessId));

    private static string PairKey(ExplorationActionCandidate candidate) =>
        PairKey(
            candidate.StateFingerprint,
            candidate.ActionKey,
            candidate.TargetHwnd,
            candidate.TargetProcessId);

    private static string PairKey(
        string stateFingerprint,
        string actionKey,
        long? targetHwnd,
        int? targetProcessId) =>
        stateFingerprint
        + "\u001f" + actionKey
        + "\u001f" + (targetHwnd is long hwnd && hwnd > 0 ? hwnd.ToString() : "-")
        + "\u001f" + (targetProcessId is int pid && pid > 0 ? pid.ToString() : "-");

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

            var byAction = string.CompareOrdinal(cx.ActionKey, cy.ActionKey);
            if (byAction != 0)
                return byAction;

            // PairKey includes StateFingerprint — comparer must too or SortedSet.Add
            // collapses distinct state/action pairs with identical action identity.
            var byFingerprint = string.CompareOrdinal(cx.StateFingerprint, cy.StateFingerprint);
            if (byFingerprint != 0)
                return byFingerprint;

            var byHwnd = Nullable.Compare(cx.TargetHwnd, cy.TargetHwnd);
            if (byHwnd != 0)
                return byHwnd;

            return Nullable.Compare(cx.TargetProcessId, cy.TargetProcessId);
        }
    }
}
