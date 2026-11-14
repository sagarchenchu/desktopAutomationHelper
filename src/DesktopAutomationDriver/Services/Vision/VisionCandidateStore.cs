using System.Collections.Concurrent;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Server-issued vision candidates. Opaque IDs + approval state; clients cannot forge authorization input.
/// </summary>
public sealed class VisionCandidateStore
{
    private readonly ConcurrentDictionary<string, VisionIssuedCandidate> _byId = new(StringComparer.Ordinal);

    public VisionIssuedCandidate Issue(VisionCandidateTarget candidate, DateTimeOffset issuedAtUtc)
    {
        ArgumentNullException.ThrowIfNull(candidate);
        if (candidate.AutoPromoted)
            throw new InvalidOperationException("Auto-promoted candidates cannot be issued.");

        var id = "vc-" + Guid.NewGuid().ToString("N");
        var issued = new VisionIssuedCandidate
        {
            Id = id,
            Snapshot = candidate with
            {
                Id = id,
                ProducedAtUtc = issuedAtUtc,
                AutoPromoted = false
            },
            IssuedAtUtc = issuedAtUtc,
            Approved = false,
            RequiresHumanApproval = true,
            AutoPromoted = false
        };

        if (!_byId.TryAdd(id, issued))
            throw new InvalidOperationException("Failed to register vision candidate id.");

        return issued;
    }

    public bool TryGet(string candidateId, out VisionIssuedCandidate issued) =>
        _byId.TryGetValue(candidateId ?? string.Empty, out issued!);

    public bool TryApprove(string candidateId, out VisionIssuedCandidate? issued)
    {
        issued = null;
        if (!_byId.TryGetValue(candidateId ?? string.Empty, out var existing))
            return false;
        if (existing.AutoPromoted)
            return false;

        existing.Approved = true;
        issued = existing;
        return true;
    }

    public void Clear() => _byId.Clear();
}

/// <summary>Server-owned vision candidate registration.</summary>
public sealed class VisionIssuedCandidate
{
    public required string Id { get; init; }
    public required VisionCandidateTarget Snapshot { get; init; }
    public DateTimeOffset IssuedAtUtc { get; init; }
    public bool Approved { get; set; }
    public bool RequiresHumanApproval { get; init; } = true;
    public bool AutoPromoted { get; init; }
}
