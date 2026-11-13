namespace DesktopAutomationAgent.Configuration;

/// <summary>
/// Bounded execution recovery budgets. Recovery never rewrites approved plans/repos.
/// </summary>
public sealed class RecoveryOptions
{
    public const string SectionName = "Recovery";

    /// <summary>When false, Phase 2 fail-fast behavior is unchanged.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>Maximum recovery attempts across the entire run.</summary>
    public int MaxTotalRecoveriesPerRun { get; set; } = 5;

    /// <summary>Maximum recovery attempts for a single failing step.</summary>
    public int MaxRecoveryAttemptsPerStep { get; set; } = 3;

    /// <summary>Additional wall-clock budget for all recovery work in a run.</summary>
    public int MaxAdditionalDurationSeconds { get; set; } = 60;

    /// <summary>Single bounded wait used by ladder tier 1 (fresh observation + wait).</summary>
    public int BoundedWaitMilliseconds { get; set; } = 750;

    /// <summary>Provider for candidate AI repair proposals. CI must use fake.</summary>
    public string Provider { get; set; } = "fake";
}
