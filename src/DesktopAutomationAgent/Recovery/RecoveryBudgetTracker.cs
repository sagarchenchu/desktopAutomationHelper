using System.Diagnostics;
using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Enforces non-recursive budgets per probe/retry/model call with accumulated recovery duration.
/// </summary>
public sealed class RecoveryBudgetTracker : IDisposable
{
    private readonly RecoveryOptions _options;
    private readonly Stopwatch _elapsed = new();
    private readonly Dictionary<string, int> _perStep = new(StringComparer.Ordinal);
    private readonly object _gate = new();
    private CancellationTokenSource? _deadlineCts;
    private int _total;
    private bool _inRecovery;
    private bool _disposed;

    public RecoveryBudgetTracker(RecoveryOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    public int TotalAttempts
    {
        get { lock (_gate) return _total; }
    }

    public bool InRecovery
    {
        get { lock (_gate) return _inRecovery; }
    }

    public TimeSpan Elapsed
    {
        get { lock (_gate) return _elapsed.Elapsed; }
    }

    /// <summary>
    /// Consumes one recovery work unit (probe, step retry, or model call).
    /// Returns a token that cancels when remaining recovery duration expires.
    /// </summary>
    public bool TryConsumeUnit(string stepId, CancellationToken parentToken, out CancellationToken recoveryToken, out string denyReason)
    {
        recoveryToken = parentToken;
        denyReason = string.Empty;

        lock (_gate)
        {
            if (_disposed)
            {
                denyReason = "recovery-disposed";
                return false;
            }

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

            if (!_elapsed.IsRunning)
                _elapsed.Start();

            if (_elapsed.Elapsed >= TimeSpan.FromSeconds(Math.Max(0, _options.MaxAdditionalDurationSeconds)))
            {
                denyReason = "max-additional-duration";
                return false;
            }

            _inRecovery = true;
            _total++;
            _perStep[stepId] = stepCount + 1;

            var remaining = TimeSpan.FromSeconds(Math.Max(0, _options.MaxAdditionalDurationSeconds)) - _elapsed.Elapsed;
            if (remaining <= TimeSpan.Zero)
            {
                _inRecovery = false;
                denyReason = "max-additional-duration";
                return false;
            }

            _deadlineCts?.Dispose();
            _deadlineCts = CancellationTokenSource.CreateLinkedTokenSource(parentToken);
            if (remaining < TimeSpan.FromMilliseconds(int.MaxValue))
                _deadlineCts.CancelAfter(remaining);
            recoveryToken = _deadlineCts.Token;
            return true;
        }
    }

    public void EndUnit()
    {
        lock (_gate)
        {
            _inRecovery = false;
        }
    }

    public TimeSpan RemainingDuration()
    {
        lock (_gate)
        {
            var budget = TimeSpan.FromSeconds(Math.Max(0, _options.MaxAdditionalDurationSeconds));
            if (!_elapsed.IsRunning)
                return budget;

            var remaining = budget - _elapsed.Elapsed;
            return remaining < TimeSpan.Zero ? TimeSpan.Zero : remaining;
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed)
                return;
            _disposed = true;
            _deadlineCts?.Dispose();
            _deadlineCts = null;
        }
    }
}
