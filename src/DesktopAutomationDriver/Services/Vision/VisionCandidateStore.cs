using System.Collections.Concurrent;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>
/// Server-issued vision candidates with TTL/capacity, action-bound approval,
/// and immutable one-use execution token grants (atomic claim before authorize).
/// </summary>
public sealed class VisionCandidateStore
{
    private readonly ConcurrentDictionary<string, VisionIssuedCandidate> _byId = new(StringComparer.Ordinal);
    /// <summary>Immutable token → grant. Entries are removed atomically on claim (one-use).</summary>
    private readonly ConcurrentDictionary<string, VisionExecutionGrant> _tokenGrants = new(StringComparer.Ordinal);
    private readonly VisionFallbackOptions _options;
    private readonly object _gate = new();

    public VisionCandidateStore(VisionFallbackOptions? options = null)
    {
        _options = options ?? new VisionFallbackOptions();
    }

    public VisionIssuedCandidate Issue(
        VisionCandidateTarget candidate,
        DateTimeOffset issuedAtUtc,
        Point? cursorScreenPoint = null)
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
            CursorScreenPoint = cursorScreenPoint,
            AutoPromoted = false,
            // Operator confirmation UX only — not OS-enforced human approval.
            RequiresHumanApproval = false
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
    /// Bind operator approval: removes any prior token for this candidate (no orphan maps),
    /// then inserts an immutable token→(candidateId, action) grant.
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
        if (string.IsNullOrWhiteSpace(executionToken))
            return false;

        lock (_gate)
        {
            if (!_byId.TryGetValue(candidateId ?? string.Empty, out var existing))
                return false;
            if (existing.AutoPromoted || existing.Consumed)
                return false;
            if (existing.ExpiresAtUtc < approvedAtUtc)
                return false;

            // Drop previous token grant so re-approval invalidates the old token (no orphans).
            if (!string.IsNullOrEmpty(existing.ExecutionToken))
                _tokenGrants.TryRemove(existing.ExecutionToken, out _);

            var normalized = NormalizeAction(action);
            var expires = approvedAtUtc.AddMilliseconds(Math.Max(1, _options.ExecutionTokenTtlMilliseconds));
            var grant = new VisionExecutionGrant(
                CandidateId: existing.Id,
                Action: normalized,
                ExpiresAtUtc: expires,
                IssuedToken: executionToken);

            if (!_tokenGrants.TryAdd(executionToken, grant))
                return false;

            existing.ApprovedAction = normalized;
            existing.ApprovedAtUtc = approvedAtUtc;
            existing.ExecutionToken = executionToken;
            existing.ExecutionTokenExpiresAtUtc = expires;
            existing.ExecutionTokenConsumed = false;
            issued = existing;
            return true;
        }
    }

    /// <summary>Peek an immutable grant without consuming (for pre-claim validation).</summary>
    public bool TryGetExecutionGrant(string executionToken, out VisionExecutionGrant grant)
    {
        grant = default!;
        EvictExpired(DateTimeOffset.UtcNow);
        if (string.IsNullOrWhiteSpace(executionToken))
            return false;
        if (!_tokenGrants.TryGetValue(executionToken, out grant!))
            return false;
        if (grant.ExpiresAtUtc < DateTimeOffset.UtcNow)
            return false;
        return true;
    }

    /// <summary>
    /// Atomically claim (remove) a one-use execution token grant. Must succeed before Allowed=true.
    /// Concurrent callers: only one TryRemove wins.
    /// </summary>
    public bool TryClaimExecutionToken(string executionToken, out VisionExecutionGrant grant)
    {
        grant = default!;
        EvictExpired(DateTimeOffset.UtcNow);
        if (string.IsNullOrWhiteSpace(executionToken))
            return false;

        if (!_tokenGrants.TryRemove(executionToken, out grant!))
            return false;

        if (grant.ExpiresAtUtc < DateTimeOffset.UtcNow)
            return false;

        if (!_byId.TryGetValue(grant.CandidateId, out var candidate))
            return false;

        // Verify submitted token still matches the issued token on the candidate.
        if (!string.Equals(candidate.ExecutionToken, executionToken, StringComparison.Ordinal)
            || !string.Equals(candidate.ExecutionToken, grant.IssuedToken, StringComparison.Ordinal))
        {
            return false;
        }

        if (candidate.Consumed || candidate.ExecutionTokenConsumed)
            return false;

        if (!string.Equals(candidate.ApprovedAction, grant.Action, StringComparison.Ordinal))
            return false;

        return true;
    }

    /// <summary>Mark candidate consumed after a successful authorize that already claimed the token.</summary>
    public bool TryConsumeAfterClaim(string candidateId, string executionToken)
    {
        if (!_byId.TryGetValue(candidateId ?? string.Empty, out var existing))
            return false;
        if (!string.Equals(existing.ExecutionToken, executionToken, StringComparison.Ordinal))
            return false;
        if (existing.Consumed)
            return false;

        existing.ExecutionTokenConsumed = true;
        existing.Consumed = true;
        // Token already removed by TryClaimExecutionToken — ensure no orphan remains.
        _tokenGrants.TryRemove(executionToken, out _);
        return true;
    }

    public void Clear()
    {
        _byId.Clear();
        _tokenGrants.Clear();
    }

    public int Count
    {
        get { EvictExpired(DateTimeOffset.UtcNow); return _byId.Count; }
    }

    public int TokenGrantCount => _tokenGrants.Count;

    public void EvictExpired(DateTimeOffset nowUtc)
    {
        foreach (var pair in _tokenGrants)
        {
            if (pair.Value.ExpiresAtUtc < nowUtc)
                _tokenGrants.TryRemove(pair.Key, out _);
        }

        foreach (var pair in _byId)
        {
            var c = pair.Value;
            if (c.ExpiresAtUtc < nowUtc || c.Consumed)
            {
                if (_byId.TryRemove(pair.Key, out var removed)
                    && !string.IsNullOrEmpty(removed.ExecutionToken))
                {
                    _tokenGrants.TryRemove(removed.ExecutionToken, out _);
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
            _tokenGrants.TryRemove(removed.ExecutionToken, out _);
        }
    }

    internal static string NormalizeAction(string? action) =>
        string.IsNullOrWhiteSpace(action) ? "click" : action.Trim().ToLowerInvariant();
}

/// <summary>Immutable one-use execution grant keyed by opaque token.</summary>
public sealed record VisionExecutionGrant(
    string CandidateId,
    string Action,
    DateTimeOffset ExpiresAtUtc,
    string IssuedToken);

/// <summary>Server-owned vision candidate registration.</summary>
public sealed class VisionIssuedCandidate
{
    public required string Id { get; init; }
    public required VisionCandidateTarget Snapshot { get; set; }
    public DateTimeOffset IssuedAtUtc { get; init; }
    public DateTimeOffset ExpiresAtUtc { get; init; }
    public Point? CursorScreenPoint { get; init; }
    /// <summary>Legacy field; always false for Phase 8 UX confirmation path.</summary>
    public bool RequiresHumanApproval { get; init; }
    public bool AutoPromoted { get; init; }

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
