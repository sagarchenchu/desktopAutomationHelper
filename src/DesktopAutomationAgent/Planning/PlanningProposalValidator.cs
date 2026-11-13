using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAutomationAgent.Driver;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.ObjectRepository;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Strict proposal schema + approved-ref + ambiguity checks. Rejects malformed,
/// extra, unknown, prohibited, ambiguous, and unsupported proposals.
/// </summary>
public sealed class PlanningProposalValidator
{
    private static readonly Regex IdPattern = new(
        @"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly HashSet<string> AllowedRisk = new(StringComparer.OrdinalIgnoreCase)
    {
        "low", "medium", "high", "prohibited"
    };

    public PlanningProposalParseResult ParseAndValidateSchema(string rawJson, int maxBytes)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(rawJson))
        {
            errors.Add("model output is empty.");
            return Fail(errors);
        }

        var redacted = SecretRedactor.Redact(rawJson);
        var byteCount = Encoding.UTF8.GetByteCount(redacted);
        if (byteCount > maxBytes)
        {
            errors.Add($"model output exceeds MaxOutputBytes ({byteCount} > {maxBytes}).");
            return Fail(errors);
        }

        PlanProposalDocument? proposal;
        try
        {
            proposal = JsonSerializer.Deserialize<PlanProposalDocument>(redacted, PlanningJson.Compact);
        }
        catch (JsonException ex)
        {
            errors.Add($"malformed JSON: {ex.Message}");
            return Fail(errors);
        }

        if (proposal is null)
        {
            errors.Add("model output deserialized to null.");
            return Fail(errors);
        }

        if (proposal.ExtensionData is { Count: > 0 })
        {
            foreach (var key in proposal.ExtensionData.Keys.OrderBy(k => k, StringComparer.Ordinal))
                errors.Add($"unknown top-level property '{key}'.");
        }

        if (proposal.SchemaVersion != PlanningVersions.ProposalSchemaVersion)
            errors.Add($"unsupported schemaVersion {proposal.SchemaVersion}; expected {PlanningVersions.ProposalSchemaVersion}.");

        if (string.IsNullOrWhiteSpace(proposal.ProposalId) || !IdPattern.IsMatch(proposal.ProposalId))
            errors.Add("proposalId is required and must match ^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$.");

        if (string.IsNullOrWhiteSpace(proposal.Goal))
            errors.Add("goal is required.");

        if (proposal.Assumptions.Count > PlanningVersions.MaxAssumptions)
            errors.Add($"assumptions exceed maximum {PlanningVersions.MaxAssumptions}.");

        if (proposal.UnresolvedQuestions.Count > PlanningVersions.MaxUnresolvedQuestions)
            errors.Add($"unresolvedQuestions exceed maximum {PlanningVersions.MaxUnresolvedQuestions}.");

        if (proposal.Steps is null || proposal.Steps.Count == 0)
            errors.Add("steps must contain at least one step.");
        else if (proposal.Steps.Count > PlanningVersions.MaxProposalSteps)
            errors.Add($"steps exceed maximum {PlanningVersions.MaxProposalSteps}.");

        if (proposal.ObjectRefs.Count > PlanningVersions.MaxRefs)
            errors.Add($"objectRefs exceed maximum {PlanningVersions.MaxRefs}.");

        if (proposal.KnowledgeRefs.Count > PlanningVersions.MaxRefs)
            errors.Add($"knowledgeRefs exceed maximum {PlanningVersions.MaxRefs}.");

        if (!AllowedRisk.Contains(proposal.Risk ?? string.Empty))
            errors.Add($"risk '{proposal.Risk}' is not allowed.");

        if (proposal.Confidence is < 0 or > 1)
            errors.Add("confidence must be between 0 and 1 when present (metadata only).");

        var seenStepIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (proposal.Steps is not null)
        {
            for (var i = 0; i < proposal.Steps.Count; i++)
            {
                var step = proposal.Steps[i];
                var loc = $"steps[{i}]";
                if (step.ExtensionData is { Count: > 0 })
                {
                    foreach (var key in step.ExtensionData.Keys)
                        errors.Add($"{loc}: unknown property '{key}'.");
                }

                if (string.IsNullOrWhiteSpace(step.Id) || !IdPattern.IsMatch(step.Id))
                    errors.Add($"{loc}: id is invalid.");
                else if (!seenStepIds.Add(step.Id))
                    errors.Add($"{loc}: duplicate step id '{step.Id}'.");

                if (string.IsNullOrWhiteSpace(step.Operation))
                    errors.Add($"{loc}: operation is required.");
                else if (PlanningPromptGuard.RequestsProhibitedCapability(step.Operation))
                    errors.Add($"{loc}: prohibited operation '{step.Operation}'.");
            }
        }

        foreach (var objectRef in proposal.ObjectRefs)
        {
            if (objectRef.ExtensionData is { Count: > 0 })
                errors.Add($"objectRefs: unknown properties on ref '{objectRef.Ref}'.");
            if (string.IsNullOrWhiteSpace(objectRef.Ref) || !objectRef.Ref.Contains('.', StringComparison.Ordinal))
                errors.Add($"objectRefs: invalid ref '{objectRef.Ref}'.");
        }

        foreach (var knowledgeRef in proposal.KnowledgeRefs)
        {
            if (knowledgeRef.ExtensionData is { Count: > 0 })
                errors.Add($"knowledgeRefs: unknown properties on '{knowledgeRef.Kind}:{knowledgeRef.Id}'.");
            if (string.IsNullOrWhiteSpace(knowledgeRef.Kind) || string.IsNullOrWhiteSpace(knowledgeRef.Id))
                errors.Add("knowledgeRefs: kind and id are required.");
        }

        return new PlanningProposalParseResult
        {
            Proposal = proposal,
            Errors = errors,
            ProposalHash = errors.Count == 0 ? Sha256Hex(PlanningJson.CanonicalSerialize(proposal)) : null
        };
    }

    public List<string> ValidateApprovedRefs(
        PlanProposalDocument proposal,
        RetrievalManifest retrieval,
        ObjectRepositorySnapshot? objectRepository,
        OperationsCatalogDto catalog)
    {
        var errors = new List<string>();
        var approvedKnowledge = retrieval.Hits
            .Where(h => h.Corpus == "knowledge")
            .Select(h => (h.Kind, h.Id))
            .ToHashSet();

        var approvedObjects = retrieval.Hits
            .Where(h => h.Corpus == "object-repository")
            .Select(h => h.Id)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (objectRepository is not null)
        {
            foreach (var (pageId, page) in objectRepository.Pages)
            {
                if (page.Elements is null)
                    continue;
                foreach (var elementId in page.Elements.Keys)
                    approvedObjects.Add($"{pageId}.{elementId}");
            }
        }

        var approvedOps = catalog.Operations
            .Select(o => o.Name)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var aliasOwners = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var op in catalog.Operations)
        {
            foreach (var alias in op.Aliases.Concat(op.DeprecatedAliases))
            {
                if (!string.IsNullOrWhiteSpace(alias))
                    aliasOwners[alias] = op.Name;
            }
        }

        foreach (var knowledgeRef in proposal.KnowledgeRefs)
        {
            if (!approvedKnowledge.Contains((knowledgeRef.Kind, knowledgeRef.Id))
                && !retrieval.Hits.Any(h =>
                    h.Corpus == "knowledge"
                    && string.Equals(h.Id, knowledgeRef.Id, StringComparison.Ordinal)))
            {
                errors.Add($"hallucinated knowledge ref '{knowledgeRef.Kind}:{knowledgeRef.Id}' is not in retrieval manifest.");
            }
        }

        foreach (var objectRef in proposal.ObjectRefs)
        {
            if (!approvedObjects.Contains(objectRef.Ref))
                errors.Add($"hallucinated object ref '{objectRef.Ref}' is not approved.");
        }

        // Embedded $objectRef in arguments
        foreach (var step in proposal.Steps)
        {
            CollectObjectRefs(step.Arguments, out var embedded);
            foreach (var embeddedRef in embedded)
            {
                if (!approvedObjects.Contains(embeddedRef))
                    errors.Add($"step '{step.Id}': embedded $objectRef '{embeddedRef}' is not approved.");
            }

            if (aliasOwners.TryGetValue(step.Operation, out var canonical))
            {
                errors.Add($"step '{step.Id}': operation '{step.Operation}' is an alias; use '{canonical}'.");
                continue;
            }

            if (!approvedOps.Contains(step.Operation))
                errors.Add($"step '{step.Id}': unsupported or unknown operation '{step.Operation}'.");
        }

        if (proposal.UnresolvedQuestions.Count > 0)
            errors.Add("proposal has unresolved questions and cannot compile.");

        return errors;
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

    private static PlanningProposalParseResult Fail(List<string> errors) =>
        new() { Errors = errors };

    private static string Sha256Hex(string text) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();
}

public sealed class PlanningProposalParseResult
{
    public PlanProposalDocument? Proposal { get; init; }

    public List<string> Errors { get; init; } = [];

    public string? ProposalHash { get; init; }

    public bool IsValid => Errors.Count == 0 && Proposal is not null;
}
