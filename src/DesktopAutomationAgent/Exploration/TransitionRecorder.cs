namespace DesktopAutomationAgent.Exploration;

public sealed class TransitionRecorder
{
    private readonly List<ExplorationStateRecord> _states = [];
    private readonly List<ExplorationTransitionRecord> _transitions = [];
    private readonly Dictionary<string, string> _stateIdsByFingerprint = new(StringComparer.Ordinal);
    private int _stateSequence;
    private int _transitionSequence;

    public IReadOnlyList<ExplorationStateRecord> States => _states;

    public IReadOnlyList<ExplorationTransitionRecord> Transitions => _transitions;

    public int MaxDepthReached { get; private set; }

    public string EnsureState(
        string fingerprint,
        string? snapshotId,
        string? windowTitle,
        string? windowClassName,
        bool? isModal,
        int depth,
        int nodeCount,
        DateTimeOffset capturedAtUtc)
    {
        if (_stateIdsByFingerprint.TryGetValue(fingerprint, out var existing))
            return existing;

        _stateSequence++;
        var stateId = $"s{_stateSequence:D4}";
        _stateIdsByFingerprint[fingerprint] = stateId;
        _states.Add(new ExplorationStateRecord
        {
            StateId = stateId,
            Fingerprint = fingerprint,
            SnapshotId = snapshotId,
            WindowTitle = windowTitle,
            WindowClassName = windowClassName,
            IsModal = isModal,
            Depth = depth,
            CapturedAtUtc = capturedAtUtc,
            NodeCount = nodeCount
        });
        MaxDepthReached = Math.Max(MaxDepthReached, depth);
        return stateId;
    }

    public ExplorationTransitionRecord Record(
        string fromStateId,
        string? toStateId,
        string fromFingerprint,
        string? toFingerprint,
        ExplorationActionCandidate action,
        ExplorationOutcome outcome,
        string? detail,
        DateTimeOffset startedAtUtc,
        DateTimeOffset completedAtUtc)
    {
        _transitionSequence++;
        var record = new ExplorationTransitionRecord
        {
            TransitionId = $"t{_transitionSequence:D4}",
            FromStateId = fromStateId,
            ToStateId = toStateId,
            FromFingerprint = fromFingerprint,
            ToFingerprint = toFingerprint,
            ActionKey = action.ActionKey,
            ActionKind = action.Kind.ToString(),
            Category = action.Category,
            DriverOperation = action.DriverOperation,
            ElementKey = action.ElementKey,
            AutomationId = action.AutomationId,
            Name = action.Name,
            Outcome = outcome,
            Detail = detail,
            StartedAtUtc = startedAtUtc,
            CompletedAtUtc = completedAtUtc,
            Depth = action.Depth
        };
        _transitions.Add(record);
        return record;
    }
}
