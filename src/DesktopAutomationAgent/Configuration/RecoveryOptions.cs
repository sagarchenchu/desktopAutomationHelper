namespace DesktopAutomationAgent.Configuration;

/// <summary>
/// Bounded execution recovery budgets. Recovery never rewrites approved plans/repos.
/// </summary>
public sealed class RecoveryOptions
{
    public const string SectionName = "Recovery";

    /// <summary>When false, Phase 2 fail-fast behavior is unchanged.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    /// When false (default), timed-out / uncertain mutation/invoke/input operations are never replayed.
    /// Only catalog read/query (or proven non-execution) may retry.
    /// </summary>
    public bool AllowMutationReplay { get; set; }

    /// <summary>Maximum recovery work units across the entire run (probe/retry/model each count).</summary>
    public int MaxTotalRecoveriesPerRun { get; set; } = 12;

    /// <summary>Maximum recovery work units for a single failing step.</summary>
    public int MaxRecoveryAttemptsPerStep { get; set; } = 8;

    /// <summary>Additional wall-clock budget for accumulated recovery work (not overall run start).</summary>
    public int MaxAdditionalDurationSeconds { get; set; } = 60;

    /// <summary>Single bounded wait used by ladder tier 1 (fresh observation + wait).</summary>
    public int BoundedWaitMilliseconds { get; set; } = 750;

    /// <summary>Max serialized candidate repair JSON bytes.</summary>
    public int MaxCandidateArtifactBytes { get; set; } = 65_536;

    /// <summary>Provider for candidate AI repair proposals. CI must use fake.</summary>
    public string Provider { get; set; } = "fake";
}
