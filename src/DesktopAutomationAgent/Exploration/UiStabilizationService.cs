using DesktopAutomationAgent.Observation;

namespace DesktopAutomationAgent.Exploration;

public sealed class StabilizationResult
{
    public bool Stabilized { get; init; }

    public bool TimedOut { get; init; }

    public string? LastFingerprint { get; init; }

    public SemanticSnapshot? Snapshot { get; init; }

    public string? Detail { get; init; }
}

/// <summary>
/// Polls observation until the state fingerprint is unchanged for a quiet period,
/// or until the stabilization budget expires.
/// </summary>
public sealed class UiStabilizationService
{
    private readonly ISemanticObservationService _observation;
    private readonly StateFingerprinter _fingerprinter;

    public UiStabilizationService(
        ISemanticObservationService observation,
        StateFingerprinter fingerprinter)
    {
        _observation = observation;
        _fingerprinter = fingerprinter;
    }

    public async Task<StabilizationResult> WaitAsync(
        ObservationOptions observationOptions,
        ExplorationBudgets budgets,
        Func<SemanticSnapshot, bool>? scopeValidator = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(observationOptions);
        ArgumentNullException.ThrowIfNull(budgets);

        var deadline = DateTimeOffset.UtcNow.AddSeconds(budgets.StabilizationTimeoutSeconds);
        var quiet = TimeSpan.FromMilliseconds(Math.Max(0, budgets.StabilizationQuietPeriodMilliseconds));
        string? lastFingerprint = null;
        DateTimeOffset stableSince = DateTimeOffset.MinValue;
        SemanticSnapshot? lastSnapshot = null;

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var observed = await _observation.ObserveAsync(observationOptions, cancellationToken).ConfigureAwait(false);
            if (!observed.Success || observed.Snapshot is null)
            {
                return new StabilizationResult
                {
                    Stabilized = false,
                    TimedOut = false,
                    Detail = observed.Error ?? "Observation failed during stabilization."
                };
            }

            lastSnapshot = observed.Snapshot;
            if (scopeValidator is not null && !scopeValidator(lastSnapshot))
            {
                return new StabilizationResult
                {
                    Stabilized = false,
                    TimedOut = false,
                    LastFingerprint = _fingerprinter.Fingerprint(lastSnapshot),
                    Snapshot = lastSnapshot,
                    Detail = "Scope violation detected during stabilization."
                };
            }

            var fingerprint = _fingerprinter.Fingerprint(lastSnapshot);
            if (string.Equals(fingerprint, lastFingerprint, StringComparison.Ordinal))
            {
                if (DateTimeOffset.UtcNow - stableSince >= quiet)
                {
                    return new StabilizationResult
                    {
                        Stabilized = true,
                        TimedOut = false,
                        LastFingerprint = fingerprint,
                        Snapshot = lastSnapshot
                    };
                }
            }
            else
            {
                lastFingerprint = fingerprint;
                stableSince = DateTimeOffset.UtcNow;
            }

            var remaining = deadline - DateTimeOffset.UtcNow;
            if (remaining <= TimeSpan.Zero)
                break;

            var delay = quiet <= TimeSpan.Zero
                ? TimeSpan.FromMilliseconds(50)
                : TimeSpan.FromMilliseconds(Math.Min(quiet.TotalMilliseconds, remaining.TotalMilliseconds));
            if (delay > TimeSpan.Zero)
                await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
        }

        return new StabilizationResult
        {
            Stabilized = false,
            TimedOut = true,
            LastFingerprint = lastFingerprint,
            Snapshot = lastSnapshot,
            Detail = $"UI did not stabilize within {budgets.StabilizationTimeoutSeconds}s."
        };
    }
}
