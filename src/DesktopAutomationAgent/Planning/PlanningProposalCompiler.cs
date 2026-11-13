using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DesktopAutomationAgent.Driver.Models;
using DesktopAutomationAgent.ObjectRepository;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Planning;

/// <summary>
/// Deterministically compiles a validated proposal into a Phase 2 plan manifest.
/// Does not call the driver.
/// </summary>
public sealed class PlanningProposalCompiler
{
    private readonly PlanValidator _planValidator = new();

    public PlanningCompileResult Compile(
        PlanProposalDocument proposal,
        string planningId,
        OperationsCatalogDto catalog,
        string? objectRepositoryRelativePath,
        PlanObjectReferenceExpander? expander = null,
        ObjectRepositorySnapshot? objectRepository = null)
    {
        ArgumentNullException.ThrowIfNull(proposal);
        ArgumentNullException.ThrowIfNull(catalog);

        var errors = new List<string>();
        var warnings = new List<string>();

        var plan = new PlanManifest
        {
            SchemaVersion = 1,
            CatalogSchemaVersion = 2,
            PlanId = "plan." + Sanitize(planningId),
            Name = Truncate("AI proposal: " + proposal.Goal, 200),
            Description = "Compiled from guarded AI plan proposal. Confidence is metadata only.",
            Tags = ["phase6", "ai-proposal"],
            ObjectRepository = objectRepositoryRelativePath,
            Metadata = new Dictionary<string, JsonElement>
            {
                ["plannerVersion"] = JsonSerializer.SerializeToElement(PlanningVersions.PlannerVersion),
                ["proposalId"] = JsonSerializer.SerializeToElement(proposal.ProposalId),
                ["planningId"] = JsonSerializer.SerializeToElement(planningId),
                ["risk"] = JsonSerializer.SerializeToElement(proposal.Risk)
            },
            Steps = proposal.Steps.Select(step => new PlanStep
            {
                Id = step.Id,
                Operation = step.Operation,
                Arguments = CloneArgs(step.Arguments),
                Assertions = step.Assertions,
                Sensitive = step.Sensitive,
                CaptureResponse = step.CaptureResponse
            }).ToList()
        };

        if (proposal.Confidence is not null)
            plan.Metadata["confidence"] = JsonSerializer.SerializeToElement(proposal.Confidence.Value);

        var relativePath = $"planning/{planningId}/candidate.plan.json";
        var validation = _planValidator.Validate(plan, relativePath);
        if (!validation.IsValid)
            errors.AddRange(validation.Errors);

        if (expander is not null && objectRepository is not null && !string.IsNullOrWhiteSpace(objectRepositoryRelativePath))
        {
            var expansion = expander.Expand(plan, objectRepository, relativePath);
            if (!expansion.Success)
                errors.AddRange(expansion.Errors);
        }

        if (errors.Count == 0)
        {
            var preflight = new PlanCatalogPreflight(catalog).Validate(plan, relativePath);
            if (preflight.Count > 0)
                errors.AddRange(preflight);
        }

        var planHash = errors.Count == 0
            ? Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(PlanningJson.CanonicalSerialize(plan)))).ToLowerInvariant()
            : null;

        return new PlanningCompileResult
        {
            Plan = errors.Count == 0 ? plan : null,
            Errors = errors,
            Warnings = warnings,
            PlanHash = planHash
        };
    }

    private static Dictionary<string, JsonElement>? CloneArgs(Dictionary<string, JsonElement>? args)
    {
        if (args is null)
            return null;
        return args.ToDictionary(p => p.Key, p => p.Value.Clone(), StringComparer.Ordinal);
    }

    private static string Sanitize(string value)
    {
        var chars = value.Trim().Select(ch => char.IsLetterOrDigit(ch) || ch is '.' or '-' or '_' ? ch : '-').ToArray();
        var sanitized = new string(chars);
        if (sanitized.Length == 0)
            sanitized = "planning";
        return sanitized.Length <= 64 ? sanitized : sanitized[..64];
    }

    private static string Truncate(string text, int max) =>
        text.Length <= max ? text : text[..max];
}

public sealed class PlanningCompileResult
{
    public PlanManifest? Plan { get; init; }

    public List<string> Errors { get; init; } = [];

    public List<string> Warnings { get; init; } = [];

    public string? PlanHash { get; init; }

    public bool IsSuccess => Errors.Count == 0 && Plan is not null;
}
