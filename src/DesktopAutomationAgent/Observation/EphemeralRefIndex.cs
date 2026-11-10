namespace DesktopAutomationAgent.Observation;

/// <summary>
/// In-memory ephemeral ref resolver scoped to a single snapshot.
/// Refs are opaque and must never be treated as persistent locators.
/// </summary>
public sealed class EphemeralRefIndex
{
    private readonly string _snapshotId;
    private readonly Dictionary<string, SemanticNode> _nodesByRef;

    public EphemeralRefIndex(SemanticSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        if (string.IsNullOrWhiteSpace(snapshot.SnapshotId))
            throw new ArgumentException("Snapshot id is required.", nameof(snapshot));

        _snapshotId = snapshot.SnapshotId;
        _nodesByRef = new Dictionary<string, SemanticNode>(StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var node in snapshot.Nodes)
        {
            if (string.IsNullOrWhiteSpace(node.Ref))
                throw new InvalidOperationException("Semantic node is missing an ephemeral ref.");

            if (!seen.Add(node.Ref))
                throw new InvalidOperationException($"Duplicate ephemeral ref '{node.Ref}'.");

            _nodesByRef[node.Ref] = node;
        }
    }

    public string SnapshotId => _snapshotId;

    public SemanticNode Resolve(string snapshotId, string nodeRef)
    {
        if (!string.Equals(snapshotId, _snapshotId, StringComparison.Ordinal))
        {
            throw new StaleEphemeralRefException(
                $"Ephemeral ref '{nodeRef}' belongs to snapshot '{snapshotId}', but resolver is bound to '{_snapshotId}'.");
        }

        if (string.IsNullOrWhiteSpace(nodeRef))
            throw new StaleEphemeralRefException("Ephemeral ref is empty or missing.");

        if (!_nodesByRef.TryGetValue(nodeRef, out var node))
            throw new StaleEphemeralRefException($"Unknown or stale ephemeral ref '{nodeRef}' for snapshot '{_snapshotId}'.");

        return node;
    }

    public bool TryResolve(string snapshotId, string nodeRef, out SemanticNode? node)
    {
        node = null;
        try
        {
            node = Resolve(snapshotId, nodeRef);
            return true;
        }
        catch (StaleEphemeralRefException)
        {
            return false;
        }
    }
}

public sealed class StaleEphemeralRefException : Exception
{
    public StaleEphemeralRefException(string message) : base(message)
    {
    }
}
