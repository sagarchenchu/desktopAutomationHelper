using DesktopAutomationAgent.ObjectRepository;

namespace DesktopAutomationAgent.Recovery;

/// <summary>
/// Deterministic fake recovery model for CI. Never contacts a provider or mutates approved repos.
/// Proposals always require human approval and are never auto-promoted.
/// </summary>
public sealed class FakeRecoveryModelClient : IRecoveryModelClient
{
    public string ProviderName => "fake";

    public Task<CandidateRepairProposal?> ProposeRepairAsync(
        RecoveryRepairRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        cancellationToken.ThrowIfCancellationRequested();

        var evidence = request.FailureEvidence ?? string.Empty;
        if (ContainsAny(evidence, "ambiguous", "assertion", "destructive", "build mismatch", "unexpected external"))
        {
            // Non-recoverable paths must not produce actionable proposals that execute.
            return Task.FromResult<CandidateRepairProposal?>(null);
        }

        var proposal = new CandidateRepairProposal
        {
            ProposalId = "rrp_" + Guid.NewGuid().ToString("N")[..12],
            StepId = request.StepId,
            Evidence = evidence,
            Confidence = 0.42,
            OldLocator = new ObjectLocator { AutomationId = "PrimaryControl" },
            NewLocator = new ObjectLocator { AutomationId = "AlternateControl", ControlType = "Button" },
            ApplicationFingerprint = request.ApplicationFingerprint,
            VerificationCommand =
                $"dotnet run --project src/DesktopAutomationAgent -- verify-object-repository --ref {request.ObjectReference ?? "page.element"}",
            AutoPromoted = false,
            RequiresHumanApproval = true,
            ObjectReference = request.ObjectReference
        };

        return Task.FromResult<CandidateRepairProposal?>(proposal);
    }

    private static bool ContainsAny(string message, params string[] needles)
    {
        foreach (var needle in needles)
        {
            if (message.Contains(needle, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return false;
    }
}
