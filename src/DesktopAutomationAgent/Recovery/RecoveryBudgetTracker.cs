using System.Diagnostics;
using DesktopAutomationAgent.Configuration;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Enforces non-recursive budgets per probe/retry/model call with accumulated
/// <em>active recovery</em> duration only (plan execution between units does not burn budget).
/// </summary>
public sealed class RecoveryBudgetTracker : IDisposable
{
    private readonly RecoveryOptions _options;
    private readonly Stopwatch _unitWatch = new();
    private readonly Dictionary<string, int> _perStep = new(StringComparer.Ordinal);
    private readonly object _gate = new();
    private CancellationTokenSource? _deadlineCts;
    private TimeSpan _accumulated;
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

    /// <summary>Accumulated active recovery time (excludes normal plan execution between units).</summary>
    public TimeSpan Elapsed
    {
        get
        {
            lock (_gate)
            {
                var total = _accumulated;
                if (_inRecovery)
                    total += _unitWatch.Elapsed;
                return total;
            }
        }
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

            var budget = TimeSpan.FromSeconds(Math.Max(0, _options.MaxAdditionalDurationSeconds));
            if (_accumulated >= budget)
            {
                denyReason = "max-additional-duration";
                return false;
            }

            var remaining = budget - _accumulated;
            if (remaining <= TimeSpan.Zero)
            {
                denyReason = "max-additional-duration";
                return false;
            }

            _inRecovery = true;
            _total++;
            _perStep[stepId] = stepCount + 1;
            _unitWatch.Restart();

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
            if (!_inRecovery)
                return;

            _unitWatch.Stop();
            _accumulated += _unitWatch.Elapsed;
            _inRecovery = false;
            _deadlineCts?.Dispose();
            _deadlineCts = null;
        }
    }

    public TimeSpan RemainingDuration()
    {
        lock (_gate)
        {
            var budget = TimeSpan.FromSeconds(Math.Max(0, _options.MaxAdditionalDurationSeconds));
            var used = _accumulated;
            if (_inRecovery)
                used += _unitWatch.Elapsed;
            var remaining = budget - used;
            return remaining < TimeSpan.Zero ? TimeSpan.Zero : remaining;
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed)
                return;
            if (_inRecovery)
            {
                _unitWatch.Stop();
                _accumulated += _unitWatch.Elapsed;
                _inRecovery = false;
            }

            _disposed = true;
            _deadlineCts?.Dispose();
            _deadlineCts = null;
        }
    }
}
