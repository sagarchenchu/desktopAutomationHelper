namespace DesktopAutomationAgent.Recovery;

public interface IRecoveryModelClient
{
    string ProviderName { get; }

    Task<CandidateRepairProposal?> ProposeRepairAsync(
        RecoveryRepairRequest request,
        CancellationToken cancellationToken = default);
}

public sealed class RecoveryRepairRequest
{
    public required string StepId { get; init; }

    public required string Operation { get; init; }

    public required string FailureEvidence { get; init; }

    public required RecoveryClassification Classification { get; init; }

    public string? ObjectReference { get; init; }

    public string? ApplicationFingerprint { get; init; }

    public string? ObservationSummary { get; init; }
}
