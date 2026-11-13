using System.Text.Json;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.ObjectRepository;

namespace DesktopAutomationAgent.Recovery;

public enum RecoveryTier
{
    None = 0,
    FreshObservationAndBoundedWait = 1,
    ReResolveApprovedLocator = 2,
    ApprovedAlternateLocator = 3,
    ApprovedEquivalentTransition = 4,
    CandidateAiRepairProposal = 5,
    FailAndCleanup = 6
}

public sealed class RecoveryAttemptRecord
{
    public int AttemptNumber { get; init; }

    public required string StepId { get; init; }

    public required RecoveryClassification Classification { get; init; }

    public required RecoveryTier Tier { get; init; }

    public required string Outcome { get; init; }

    public string? Detail { get; init; }

    public string? SnapshotBeforePath { get; init; }

    public string? SnapshotAfterPath { get; init; }

    public string? CandidateRepairPath { get; init; }

    public DateTimeOffset StartedAtUtc { get; init; }

    public DateTimeOffset FinishedAtUtc { get; init; }

    public double DurationMilliseconds { get; init; }
}

public sealed class RecoveryRunSummary
{
    public required string RunId { get; init; }

    public required string OriginalPlanSha256 { get; init; }

    public string? ObjectRepositorySha256 { get; init; }

    public string? ObjectRepositorySha256After { get; init; }

    public bool PlanMutated { get; init; }

    public bool ApprovedRepositoryMutated { get; init; }

    public int TotalRecoveryAttempts { get; init; }

    public IReadOnlyList<RecoveryAttemptRecord> Attempts { get; init; } = Array.Empty<RecoveryAttemptRecord>();

    public string FinalResult { get; init; } = "none";
}

public sealed class RecoveryContext
{
    public required string RunId { get; init; }

    public required string RunDirectory { get; init; }

    public required string PlanPath { get; init; }

    public required string OriginalPlanSha256 { get; init; }

    public string? ObjectRepositoryPath { get; init; }

    public string? ObjectRepositorySha256 { get; init; }

    public ObjectRepositorySnapshot? RepositorySnapshot { get; init; }

    public OperationsCatalogDto? Catalog { get; init; }

    public IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> StepObjectBindings { get; init; }
        = new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.Ordinal);

    public IReadOnlyDictionary<string, ApprovedEquivalentTransition>? EquivalentTransitions { get; init; }

    /// <summary>objectRef → approved candidate locator ready for unique verification.</summary>
    public IReadOnlyDictionary<string, ObjectLocator>? ApprovedCandidateLocators { get; init; }

    public IReadOnlySet<string>? ApprovedCandidateLocatorIds { get; init; }

    public string? ApplicationFingerprint { get; init; }

    public string? ExpectedApplicationFingerprint { get; init; }
}

public sealed class ApprovedEquivalentTransition
{
    public required string TransitionId { get; init; }

    public required string FromAction { get; init; }

    public required string EquivalentAction { get; init; }

    public Dictionary<string, JsonElement>? ReplacementArguments { get; init; }

    public string? KnowledgeRef { get; init; }
}

public sealed class CandidateRepairProposal
{
    public required string ProposalId { get; init; }

    public required string StepId { get; init; }

    public required string Evidence { get; init; }

    public double Confidence { get; init; }

    public ObjectLocator? OldLocator { get; init; }

    public ObjectLocator? NewLocator { get; init; }

    public string? ApplicationFingerprint { get; init; }

    public required string VerificationCommand { get; init; }

    public bool AutoPromoted { get; set; }

    public bool RequiresHumanApproval { get; set; } = true;

    public string? ObjectReference { get; init; }
}

public sealed class RecoveryDecision
{
    public required RecoveryClassification Classification { get; init; }

    public required bool MayAttemptRecovery { get; init; }

    public string? Evidence { get; init; }

    public string? ObjectReference { get; init; }

    /// <summary>Exact locator argument key that failed (locator/locator2/parentLocator/containerLocator).</summary>
    public string? BindingKey { get; init; }
}

public sealed class RecoveryOutcome
{
    public bool Recovered { get; init; }

    public bool StopForApproval { get; init; }

    public RecoveryTier TierUsed { get; init; }

    public RecoveryClassification Classification { get; init; }

    public string? Detail { get; init; }

    public string? CandidateRepairPath { get; init; }

    public Plans.PlanStep? RewrittenStep { get; init; }

    public Execution.StepRunResult? RecoveredStepResult { get; init; }

    public IReadOnlyList<RecoveryAttemptRecord> Attempts { get; init; } = Array.Empty<RecoveryAttemptRecord>();
}
