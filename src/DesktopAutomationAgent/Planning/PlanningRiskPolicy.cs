using DesktopAutomationAgent.Exploration;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Maps proposal risk + unresolved questions + requested mode to an effective approval mode.
/// Confidence never authorizes execution.
/// </summary>
public static class PlanningRiskPolicy
{
    private static readonly HashSet<string> DataEntryOps = new(StringComparer.OrdinalIgnoreCase)
    {
        "typeuia",
        "setvalue",
        "sendkeys",
        "cleartext"
    };

    private static readonly HashSet<string> TransactionalTerms = new(StringComparer.OrdinalIgnoreCase)
    {
        "submit", "save", "pay", "purchase", "send", "transfer", "approve", "publish"
    };

    public static PlanningRiskLevel ParseRisk(string? risk) =>
        risk?.Trim().ToLowerInvariant() switch
        {
            "low" => PlanningRiskLevel.Low,
            "medium" => PlanningRiskLevel.Medium,
            "high" => PlanningRiskLevel.High,
            "prohibited" => PlanningRiskLevel.Prohibited,
            _ => PlanningRiskLevel.High
        };

    public static PlanningApprovalMode Decide(
        PlanProposalDocument proposal,
        PlanningApprovalMode requestedMode,
        out List<string> reasons)
    {
        reasons = [];
        var risk = ParseRisk(proposal.Risk);

        if (risk == PlanningRiskLevel.Prohibited)
        {
            reasons.Add("risk=prohibited.");
            return PlanningApprovalMode.Denied;
        }

        if (proposal.UnresolvedQuestions.Count > 0)
        {
            reasons.Add("unresolved questions present.");
            return PlanningApprovalMode.Denied;
        }

        if (HasDestructiveSteps(proposal, out var destructiveDetail))
        {
            reasons.Add(destructiveDetail);
            return PlanningApprovalMode.Denied;
        }

        if (HasProhibitedCapability(proposal, out var prohibitedDetail))
        {
            reasons.Add(prohibitedDetail);
            return PlanningApprovalMode.Denied;
        }

        var approvalDetail = string.Empty;
        var needsApproval = risk is PlanningRiskLevel.Medium or PlanningRiskLevel.High
            || HasDataEntryOrTransaction(proposal, out approvalDetail);

        if (needsApproval)
        {
            if (!string.IsNullOrWhiteSpace(approvalDetail))
                reasons.Add(approvalDetail);
            reasons.Add($"risk={risk.ToString().ToLowerInvariant()} requires human approval.");

            // Requested denied stays denied; otherwise approval-required.
            if (requestedMode == PlanningApprovalMode.Denied)
                return PlanningApprovalMode.Denied;

            return PlanningApprovalMode.ApprovalRequired;
        }

        // Low risk, fully mapped: honor request but never upgrade to auto-run unless requested.
        return requestedMode switch
        {
            PlanningApprovalMode.AutoRunKnownSafe => PlanningApprovalMode.AutoRunKnownSafe,
            PlanningApprovalMode.Denied => PlanningApprovalMode.Denied,
            PlanningApprovalMode.ApprovalRequired => PlanningApprovalMode.ApprovalRequired,
            _ => PlanningApprovalMode.ProposalOnly
        };
    }

    public static string ToWire(PlanningApprovalMode mode) =>
        mode switch
        {
            PlanningApprovalMode.AutoRunKnownSafe => "auto-run-known-safe",
            PlanningApprovalMode.ApprovalRequired => "approval-required",
            PlanningApprovalMode.Denied => "denied",
            _ => "proposal-only"
        };

    public static bool TryParseMode(string? text, out PlanningApprovalMode mode)
    {
        mode = PlanningApprovalMode.ProposalOnly;
        if (string.IsNullOrWhiteSpace(text))
            return true;

        switch (text.Trim().ToLowerInvariant())
        {
            case "proposal-only":
            case "proposalonly":
                mode = PlanningApprovalMode.ProposalOnly;
                return true;
            case "auto-run-known-safe":
            case "autorunknownsafe":
                mode = PlanningApprovalMode.AutoRunKnownSafe;
                return true;
            case "approval-required":
            case "approvalrequired":
                mode = PlanningApprovalMode.ApprovalRequired;
                return true;
            case "denied":
                mode = PlanningApprovalMode.Denied;
                return true;
            default:
                return false;
        }
    }

    private static bool HasDestructiveSteps(PlanProposalDocument proposal, out string detail)
    {
        foreach (var step in proposal.Steps)
        {
            var blob = $"{step.Operation} {step.Rationale} {step.Id}";
            if (PlanningPromptGuard.ContainsDestructiveIntent(ForbiddenActionVocabulary.Normalize(blob))
                || ForbiddenActionVocabulary.Matches(step.Rationale) &&
                   ForbiddenActionVocabulary.Normalize(step.Rationale ?? string.Empty)
                       .Contains("delete", StringComparison.Ordinal))
            {
                detail = $"destructive step '{step.Id}'.";
                return true;
            }
        }

        detail = string.Empty;
        return false;
    }

    private static bool HasProhibitedCapability(PlanProposalDocument proposal, out string detail)
    {
        foreach (var step in proposal.Steps)
        {
            if (PlanningPromptGuard.RequestsProhibitedCapability($"{step.Operation} {step.Rationale}"))
            {
                detail = $"prohibited capability in step '{step.Id}'.";
                return true;
            }
        }

        detail = string.Empty;
        return false;
    }

    private static bool HasDataEntryOrTransaction(PlanProposalDocument proposal, out string detail)
    {
        foreach (var step in proposal.Steps)
        {
            if (DataEntryOps.Contains(step.Operation))
            {
                detail = $"data-entry operation '{step.Operation}'.";
                return true;
            }

            var text = ForbiddenActionVocabulary.Normalize($"{step.Operation} {step.Rationale} {step.Id}");
            foreach (var term in TransactionalTerms)
            {
                if (text.Contains(term, StringComparison.Ordinal))
                {
                    detail = $"transactional intent '{term}'.";
                    return true;
                }
            }
        }

        var goal = ForbiddenActionVocabulary.Normalize(proposal.Goal);
        foreach (var term in TransactionalTerms)
        {
            if (goal.Contains(term, StringComparison.Ordinal))
            {
                detail = $"transactional goal '{term}'.";
                return true;
            }
        }

        detail = string.Empty;
        return false;
    }
}
