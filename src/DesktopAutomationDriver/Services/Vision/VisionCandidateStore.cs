using System.Collections.Concurrent;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Server-issued vision candidates with TTL, capacity eviction, action-bound approval,
/// short-lived one-use execution tokens, and consume-on-authorize.
/// </summary>
public sealed class VisionCandidateStore
{
    private readonly ConcurrentDictionary<string, VisionIssuedCandidate> _byId = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, string> _tokenToCandidateId = new(StringComparer.Ordinal);
    private readonly VisionFallbackOptions _options;
    private readonly object _gate = new();

    public VisionCandidateStore(VisionFallbackOptions? options = null)
    {
        _options = options ?? new VisionFallbackOptions();
    }

    public VisionIssuedCandidate Issue(VisionCandidateTarget candidate, DateTimeOffset issuedAtUtc)
    {
        ArgumentNullException.ThrowIfNull(candidate);
        if (candidate.AutoPromoted)
            throw new InvalidOperationException("Auto-promoted candidates cannot be issued.");

        EvictExpired(issuedAtUtc);

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
            ExpiresAtUtc = issuedAtUtc.AddMilliseconds(Math.Max(1, _options.CandidateTtlMilliseconds)),
            AutoPromoted = false,
            RequiresHumanApproval = true
        };

        lock (_gate)
        {
            while (_byId.Count >= Math.Max(1, _options.MaxStoredCandidates))
                EvictOldestUnlocked();

            if (!_byId.TryAdd(id, issued))
                throw new InvalidOperationException("Failed to register vision candidate id.");
        }

        return issued;
    }

    public bool TryGet(string candidateId, out VisionIssuedCandidate issued)
    {
        issued = null!;
        EvictExpired(DateTimeOffset.UtcNow);
        return _byId.TryGetValue(candidateId ?? string.Empty, out issued!);
    }

    /// <summary>
    /// Bind operator approval to candidate + action and attach a short-lived one-use execution token.
    /// </summary>
    public bool TryBindApproval(
        string candidateId,
        string action,
        string executionToken,
        DateTimeOffset approvedAtUtc,
        out VisionIssuedCandidate? issued)
    {
        issued = null;
        EvictExpired(approvedAtUtc);
        if (!_byId.TryGetValue(candidateId ?? string.Empty, out var existing))
            return false;
        if (existing.AutoPromoted || existing.Consumed)
            return false;
        if (existing.ExpiresAtUtc < approvedAtUtc)
            return false;

        existing.ApprovedAction = NormalizeAction(action);
        existing.ApprovedAtUtc = approvedAtUtc;
        existing.ExecutionToken = executionToken;
        existing.ExecutionTokenExpiresAtUtc =
            approvedAtUtc.AddMilliseconds(Math.Max(1, _options.ExecutionTokenTtlMilliseconds));
        existing.ExecutionTokenConsumed = false;
        _tokenToCandidateId[executionToken] = existing.Id;
        issued = existing;
        return true;
    }

    public bool TryGetByExecutionToken(string executionToken, out VisionIssuedCandidate issued)
    {
        issued = null!;
        EvictExpired(DateTimeOffset.UtcNow);
        if (string.IsNullOrWhiteSpace(executionToken))
            return false;
        if (!_tokenToCandidateId.TryGetValue(executionToken, out var candidateId))
            return false;
        return _byId.TryGetValue(candidateId, out issued!);
    }

    /// <summary>Consume candidate + execution token after a successful authorize (one-use).</summary>
    public bool TryConsume(string candidateId, string executionToken)
    {
        if (!_byId.TryGetValue(candidateId ?? string.Empty, out var existing))
            return false;
        if (!string.Equals(existing.ExecutionToken, executionToken, StringComparison.Ordinal))
            return false;
        if (existing.ExecutionTokenConsumed || existing.Consumed)
            return false;

        existing.ExecutionTokenConsumed = true;
        existing.Consumed = true;
        _tokenToCandidateId.TryRemove(executionToken, out _);
        return true;
    }

    public void Clear()
    {
        _byId.Clear();
        _tokenToCandidateId.Clear();
    }

    public int Count
    {
        get { EvictExpired(DateTimeOffset.UtcNow); return _byId.Count; }
    }

    public void EvictExpired(DateTimeOffset nowUtc)
    {
        foreach (var pair in _byId)
        {
            var c = pair.Value;
            var expired = c.ExpiresAtUtc < nowUtc
                          || (c.ExecutionTokenExpiresAtUtc is { } tokenExp
                              && tokenExp < nowUtc
                              && c.ExecutionTokenConsumed);
            if (expired || c.Consumed)
            {
                if (_byId.TryRemove(pair.Key, out var removed)
                    && !string.IsNullOrEmpty(removed.ExecutionToken))
                {
                    _tokenToCandidateId.TryRemove(removed.ExecutionToken, out _);
                }
            }
        }
    }

    private void EvictOldestUnlocked()
    {
        VisionIssuedCandidate? oldest = null;
        foreach (var pair in _byId)
        {
            if (oldest is null || pair.Value.IssuedAtUtc < oldest.IssuedAtUtc)
                oldest = pair.Value;
        }

        if (oldest is null)
            return;

        if (_byId.TryRemove(oldest.Id, out var removed)
            && !string.IsNullOrEmpty(removed.ExecutionToken))
        {
            _tokenToCandidateId.TryRemove(removed.ExecutionToken, out _);
        }
    }

    internal static string NormalizeAction(string? action) =>
        string.IsNullOrWhiteSpace(action) ? "click" : action.Trim().ToLowerInvariant();
}

/// <summary>Server-owned vision candidate registration.</summary>
public sealed class VisionIssuedCandidate
{
    public required string Id { get; init; }
    public required VisionCandidateTarget Snapshot { get; set; }
    public DateTimeOffset IssuedAtUtc { get; init; }
    public DateTimeOffset ExpiresAtUtc { get; init; }
    public bool RequiresHumanApproval { get; init; } = true;
    public bool AutoPromoted { get; init; }

    /// <summary>Action bound at operator approval time (null until approved).</summary>
    public string? ApprovedAction { get; set; }
    public DateTimeOffset? ApprovedAtUtc { get; set; }
    public string? ExecutionToken { get; set; }
    public DateTimeOffset? ExecutionTokenExpiresAtUtc { get; set; }
    public bool ExecutionTokenConsumed { get; set; }
    public bool Consumed { get; set; }

    public bool IsApproved =>
        !string.IsNullOrWhiteSpace(ApprovedAction)
        && !string.IsNullOrWhiteSpace(ExecutionToken)
        && !Consumed;
}
