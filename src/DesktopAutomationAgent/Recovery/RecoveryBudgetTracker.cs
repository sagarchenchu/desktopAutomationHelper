using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Enforces non-recursive, per-run / per-step / duration recovery budgets.
/// </summary>
public sealed class RecoveryBudgetTracker
{
    private readonly RecoveryOptions _options;
    private readonly DateTimeOffset _deadline;
    private readonly Dictionary<string, int> _perStep = new(StringComparer.Ordinal);
    private int _total;
    private bool _inRecovery;

    public RecoveryBudgetTracker(RecoveryOptions options, DateTimeOffset? startedAtUtc = null)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        var start = startedAtUtc ?? DateTimeOffset.UtcNow;
        _deadline = start.AddSeconds(Math.Max(0, _options.MaxAdditionalDurationSeconds));
    }

    public int TotalAttempts => _total;

    public bool InRecovery => _inRecovery;

    public bool TryBeginAttempt(string stepId, out string denyReason)
    {
        denyReason = string.Empty;

        if (!_options.Enabled)
        {
            denyReason = "recovery-disabled";
            return false;
        }

        if (_inRecovery)
        {
            denyReason = "no-recursive-recovery";
            return false;
        }

        if (_total >= _options.MaxTotalRecoveriesPerRun)
        {
            denyReason = "max-total-recoveries";
            return false;
        }

        var stepCount = _perStep.GetValueOrDefault(stepId, 0);
        if (stepCount >= _options.MaxRecoveryAttemptsPerStep)
        {
            denyReason = "max-per-step-recoveries";
            return false;
        }

        if (DateTimeOffset.UtcNow >= _deadline)
        {
            denyReason = "max-additional-duration";
            return false;
        }

        _inRecovery = true;
        _total++;
        _perStep[stepId] = stepCount + 1;
        return true;
    }

    public void EndAttempt() => _inRecovery = false;

    public TimeSpan RemainingDuration()
    {
        var remaining = _deadline - DateTimeOffset.UtcNow;
        return remaining < TimeSpan.Zero ? TimeSpan.Zero : remaining;
    }
}
