using System.Collections.Concurrent;
using DesktopAutomationAgent.Observation;

namespace DesktopAutomationAgent.Mcp;

/// <summary>
/// In-process opaque ref lifetime for MCP clients. Refs are snapshot-scoped and expire;
/// they are never durable locators and never include runtime IDs.
/// </summary>
public sealed class McpOpaqueRefStore
{
    private readonly ConcurrentDictionary<string, RefEntry> _entries = new(StringComparer.Ordinal);
    private readonly TimeSpan _lifetime;
    private readonly object _gate = new();
    private string? _activeSnapshotId;

    public McpOpaqueRefStore(TimeSpan? lifetime = null)
    {
        _lifetime = lifetime ?? TimeSpan.FromMinutes(15);
    }

    public TimeSpan Lifetime => _lifetime;

    public string? ActiveSnapshotId
    {
        get
        {
            lock (_gate)
                return _activeSnapshotId;
        }
    }

    public void ReplaceSnapshot(SemanticSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        var expires = DateTimeOffset.UtcNow.Add(_lifetime);
        lock (_gate)
        {
            _entries.Clear();
            _activeSnapshotId = snapshot.SnapshotId;
            foreach (var node in snapshot.Nodes)
            {
                if (string.IsNullOrWhiteSpace(node.Ref))
                    continue;
                _entries[node.Ref] = new RefEntry(snapshot.SnapshotId, node, expires);
            }
        }
    }

    public SemanticNode Resolve(string snapshotId, string nodeRef)
    {
        if (string.IsNullOrWhiteSpace(snapshotId) || string.IsNullOrWhiteSpace(nodeRef))
            throw new McpToolException(McpErrorCode.StaleRef, "Ephemeral ref or snapshot id is missing.");

        if (!_entries.TryGetValue(nodeRef, out var entry))
            throw new McpToolException(McpErrorCode.StaleRef, $"Unknown or stale ephemeral ref '{nodeRef}'.");

        if (!string.Equals(entry.SnapshotId, snapshotId, StringComparison.Ordinal))
        {
            throw new McpToolException(
                McpErrorCode.StaleRef,
                $"Ephemeral ref '{nodeRef}' belongs to snapshot '{entry.SnapshotId}', not '{snapshotId}'.");
        }

        if (DateTimeOffset.UtcNow > entry.ExpiresAtUtc)
        {
            _entries.TryRemove(nodeRef, out _);
            throw new McpToolException(McpErrorCode.StaleRef, $"Ephemeral ref '{nodeRef}' has expired.");
        }

        return entry.Node;
    }

    public void ExpireAll()
    {
        lock (_gate)
        {
            _entries.Clear();
            _activeSnapshotId = null;
        }
    }

    public int Count => _entries.Count;

    private sealed record RefEntry(string SnapshotId, SemanticNode Node, DateTimeOffset ExpiresAtUtc);
}
