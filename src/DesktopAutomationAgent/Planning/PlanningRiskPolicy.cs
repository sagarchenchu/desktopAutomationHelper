using System.Text.Json;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.Exploration;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Maps proposal risk + catalog classification + retrieval constraints to an effective approval mode.
/// Confidence never authorizes execution. Auto-run-known-safe never authorizes mutations or raw locators.
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

    /// <summary>
    /// Explicit safe allowlist for auto-run (in addition to catalog query/read classification).
    /// </summary>
    public static readonly HashSet<string> ExplicitSafeReadOperations = new(StringComparer.OrdinalIgnoreCase)
    {
        "listwindows",
        "finduia",
        "dumpuia"
    };

    private static readonly HashSet<string> LocatorArgumentKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "locator",
        "locator2",
        "parentLocator",
        "containerLocator"
    };

    private static readonly HashSet<string> RawLocatorIdentityKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "automationId",
        "name",
        "className",
        "controlType",
        "frameworkId",
        "foundIndex",
        "xpath",
        "path"
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
        OperationsCatalogDto catalog,
        RetrievalManifest retrieval,
        out List<string> reasons)
    {
        ArgumentNullException.ThrowIfNull(proposal);
        ArgumentNullException.ThrowIfNull(catalog);
        ArgumentNullException.ThrowIfNull(retrieval);

        reasons = [];
        var risk = ParseRisk(proposal.Risk);

        if (risk == PlanningRiskLevel.Prohibited)
        {
            reasons.Add("risk=prohibited.");
            return PlanningApprovalMode.Denied;
        }

        if ((proposal.UnresolvedQuestions?.Count ?? 0) > 0)
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
        var mutationDetail = string.Empty;
        var needsApproval = risk is PlanningRiskLevel.Medium or PlanningRiskLevel.High
            || HasDataEntryOrTransaction(proposal, out approvalDetail)
            || HasNonSafeMutation(proposal, catalog, out mutationDetail);

        if (needsApproval)
        {
            if (!string.IsNullOrWhiteSpace(approvalDetail))
                reasons.Add(approvalDetail);
            if (!string.IsNullOrWhiteSpace(mutationDetail))
                reasons.Add(mutationDetail);
            reasons.Add($"risk={risk.ToString().ToLowerInvariant()} requires human approval.");

            if (requestedMode == PlanningApprovalMode.Denied)
                return PlanningApprovalMode.Denied;

            return PlanningApprovalMode.ApprovalRequired;
        }

        if (requestedMode == PlanningApprovalMode.AutoRunKnownSafe)
        {
            if (!TryAuthorizeAutoRun(proposal, catalog, retrieval, out var autoRunReasons))
            {
                reasons.AddRange(autoRunReasons);
                return PlanningApprovalMode.ApprovalRequired;
            }

            return PlanningApprovalMode.AutoRunKnownSafe;
        }

        return requestedMode switch
        {
            PlanningApprovalMode.Denied => PlanningApprovalMode.Denied,
            PlanningApprovalMode.ApprovalRequired => PlanningApprovalMode.ApprovalRequired,
            _ => PlanningApprovalMode.ProposalOnly
        };
    }

    /// <summary>
    /// Auto-run requires catalog read/query (or explicit safe allowlist), retrieved mappings,
    /// and forbids model-supplied raw locators.
    /// </summary>
    public static bool TryAuthorizeAutoRun(
        PlanProposalDocument proposal,
        OperationsCatalogDto catalog,
        RetrievalManifest retrieval,
        out List<string> reasons)
    {
        reasons = [];
        var byName = catalog.Operations
            .Where(o => !string.IsNullOrWhiteSpace(o.Name))
            .GroupBy(o => o.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var retrievedMappings = retrieval.Hits
            .Where(h => string.Equals(h.Corpus, "step-registry", StringComparison.Ordinal))
            .Select(h => h.Id)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var retrievedObjects = retrieval.Hits
            .Where(h => string.Equals(h.Corpus, "object-repository", StringComparison.Ordinal))
            .Select(h => h.Id)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var step in proposal.Steps ?? [])
        {
            if (!byName.TryGetValue(step.Operation, out var descriptor))
            {
                reasons.Add($"auto-run: unknown operation '{step.Operation}' on step '{step.Id}'.");
                continue;
            }

            if (!IsSafeReadOrQuery(descriptor))
            {
                reasons.Add(
                    $"auto-run: operation '{step.Operation}' is not catalog read/query " +
                    $"(category='{descriptor.Category}', operationType='{descriptor.OperationType}').");
            }

            if (string.IsNullOrWhiteSpace(step.MappingId) || !retrievedMappings.Contains(step.MappingId))
            {
                reasons.Add(
                    $"auto-run: step '{step.Id}' requires a retrieved step-registry mappingId.");
            }

            if (HasRawLocator(step, out var rawDetail))
                reasons.Add($"auto-run: {rawDetail}");

            CollectObjectRefs(step.Arguments, out var embedded);
            foreach (var objectRef in embedded)
            {
                if (!retrievedObjects.Contains(objectRef))
                {
                    reasons.Add(
                        $"auto-run: embedded object ref '{objectRef}' on step '{step.Id}' was not retrieved.");
                }
            }
        }

        foreach (var objectRef in proposal.ObjectRefs ?? [])
        {
            if (!retrievedObjects.Contains(objectRef.Ref))
            {
                reasons.Add($"auto-run: object ref '{objectRef.Ref}' was not retrieved.");
            }
        }

        return reasons.Count == 0;
    }

    public static bool IsSafeReadOrQuery(OperationDescriptorDto descriptor)
    {
        ArgumentNullException.ThrowIfNull(descriptor);
        if (!string.IsNullOrWhiteSpace(descriptor.Name) && ExplicitSafeReadOperations.Contains(descriptor.Name))
            return true;

        var category = (descriptor.Category ?? string.Empty).Trim().ToLowerInvariant();
        var operationType = (descriptor.OperationType ?? string.Empty).Trim().ToLowerInvariant();
        return category is "query" or "read" or "observation"
            || operationType is "query" or "read" or "observation";
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

    private static bool HasNonSafeMutation(
        PlanProposalDocument proposal,
        OperationsCatalogDto catalog,
        out string detail)
    {
        var byName = catalog.Operations
            .Where(o => !string.IsNullOrWhiteSpace(o.Name))
            .GroupBy(o => o.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        foreach (var step in proposal.Steps ?? [])
        {
            if (!byName.TryGetValue(step.Operation, out var descriptor))
                continue;

            if (!IsSafeReadOrQuery(descriptor))
            {
                detail =
                    $"mutation/invoke/input operation '{step.Operation}' " +
                    $"(category='{descriptor.Category}', operationType='{descriptor.OperationType}') requires approval.";
                return true;
            }
        }

        detail = string.Empty;
        return false;
    }

    public static bool HasRawLocator(PlanProposalStep step, out string detail)
    {
        detail = string.Empty;
        if (step.Arguments is null)
            return false;

        foreach (var (key, value) in step.Arguments)
        {
            if (!LocatorArgumentKeys.Contains(key))
                continue;

            if (value.ValueKind == JsonValueKind.String)
            {
                var text = value.GetString() ?? string.Empty;
                if (text.StartsWith("$objectRef:", StringComparison.Ordinal))
                    continue;

                detail = $"step '{step.Id}' argument '{key}' uses a raw locator string.";
                return true;
            }

            if (value.ValueKind != JsonValueKind.Object)
            {
                detail = $"step '{step.Id}' argument '{key}' is not an approved $objectRef locator.";
                return true;
            }

            if (!value.TryGetProperty("$objectRef", out var objectRef)
                || objectRef.ValueKind != JsonValueKind.String
                || string.IsNullOrWhiteSpace(objectRef.GetString()))
            {
                detail = $"step '{step.Id}' argument '{key}' uses a model-supplied raw locator.";
                return true;
            }

            foreach (var property in value.EnumerateObject())
            {
                if (property.NameEquals("$objectRef"))
                    continue;
                if (RawLocatorIdentityKeys.Contains(property.Name))
                {
                    detail =
                        $"step '{step.Id}' argument '{key}' mixes $objectRef with raw identity field '{property.Name}'.";
                    return true;
                }
            }
        }

        return false;
    }

    private static void CollectObjectRefs(Dictionary<string, JsonElement>? arguments, out List<string> refs)
    {
        refs = [];
        if (arguments is null)
            return;

        foreach (var value in arguments.Values)
            Walk(value, refs);
    }

    private static void Walk(JsonElement element, List<string> refs)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                if (element.TryGetProperty("$objectRef", out var objectRef)
                    && objectRef.ValueKind == JsonValueKind.String)
                {
                    var value = objectRef.GetString();
                    if (!string.IsNullOrWhiteSpace(value))
                        refs.Add(value);
                }

                foreach (var property in element.EnumerateObject())
                    Walk(property.Value, refs);
                break;
            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                    Walk(item, refs);
                break;
            case JsonValueKind.String:
                var text = element.GetString();
                if (!string.IsNullOrWhiteSpace(text)
                    && text.StartsWith("$objectRef:", StringComparison.Ordinal))
                {
                    refs.Add(text["$objectRef:".Length..]);
                }

                break;
        }
    }

    private static bool HasDestructiveSteps(PlanProposalDocument proposal, out string detail)
    {
        foreach (var step in proposal.Steps ?? [])
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
        foreach (var step in proposal.Steps ?? [])
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
        foreach (var step in proposal.Steps ?? [])
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
