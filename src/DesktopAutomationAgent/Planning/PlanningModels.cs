using System.Text.Json;
using System.Text.Json.Serialization;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Planning;

public enum PlanningApprovalMode
{
    ProposalOnly,
    AutoRunKnownSafe,
    ApprovalRequired,
    Denied
}

public enum PlanningRiskLevel
{
    Low,
    Medium,
    High,
    Prohibited
}

public sealed class PlanningRequest
{
    public string Goal { get; init; } = string.Empty;

    /// <summary>Optional untrusted UI/Jira text — never trusted for policy.</summary>
    public string? UntrustedContext { get; init; }

    public string? ApplicationId { get; init; }

    public string? ExpectedBuildFingerprint { get; init; }

    public string? KnowledgePath { get; init; }

    public string? ObjectRepositoryPath { get; init; }

    public string? StepRegistryPath { get; init; }

    public string? CatalogPath { get; init; }

    public PlanningApprovalMode RequestedApprovalMode { get; init; } = PlanningApprovalMode.ProposalOnly;

    /// <summary>
    /// When true and policy authorizes auto-run-known-safe, execute via Phase 2 runner.
    /// Proposal-only / denied / approval-required never execute even when set.
    /// </summary>
    public bool ExecuteIfAuthorized { get; init; }

    public bool DryRun { get; init; } = true;

    public string? PlanningId { get; init; }

    public CancellationToken CancellationToken { get; init; }
}

public sealed class PlanProposalDocument
{
    [JsonPropertyName("$schema")]
    public string? Schema { get; set; }

    public int SchemaVersion { get; set; } = PlanningVersions.ProposalSchemaVersion;

    public string ProposalId { get; set; } = string.Empty;

    public string Goal { get; set; } = string.Empty;

    public List<string> Assumptions { get; set; } = [];

    public List<string> UnresolvedQuestions { get; set; } = [];

    public List<PlanProposalStep> Steps { get; set; } = [];

    public List<PlanProposalObjectRef> ObjectRefs { get; set; } = [];

    public List<PlanProposalKnowledgeRef> KnowledgeRefs { get; set; } = [];

    public string? ExpectedState { get; set; }

    public string? Postcondition { get; set; }

    public string Risk { get; set; } = "low";

    /// <summary>Metadata only — never authorizes execution.</summary>
    public double? Confidence { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class PlanProposalStep
{
    public string Id { get; set; } = string.Empty;

    public string Operation { get; set; } = string.Empty;

    public Dictionary<string, JsonElement>? Arguments { get; set; }

    public List<PlanAssertion>? Assertions { get; set; }

    public bool Sensitive { get; set; }

    public bool? CaptureResponse { get; set; }

    public string? MappingId { get; set; }

    public string? Rationale { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class PlanProposalObjectRef
{
    public string Ref { get; set; } = string.Empty;

    public string Role { get; set; } = "target";

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class PlanProposalKnowledgeRef
{
    public string Kind { get; set; } = string.Empty;

    public string Id { get; set; } = string.Empty;

    public string Role { get; set; } = "navigation";

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class RetrievalHitRecord
{
    public string Corpus { get; set; } = string.Empty;

    public string Kind { get; set; } = string.Empty;

    public string Id { get; set; } = string.Empty;

    public int Score { get; set; }

    public string ContentHash { get; set; } = string.Empty;

    public string Snippet { get; set; } = string.Empty;
}

public sealed class RetrievalManifest
{
    public string RetrievalVersion { get; set; } = PlanningVersions.RetrievalVersion;

    public string CanonicalGoal { get; set; } = string.Empty;

    public string GoalHash { get; set; } = string.Empty;

    public string? ApplicationId { get; set; }

    public string? BuildFingerprint { get; set; }

    public string? KnowledgeContentHash { get; set; }

    public string? ObjectRepositoryHash { get; set; }

    public string? StepRegistryHash { get; set; }

    public string? CatalogFingerprint { get; set; }

    public List<RetrievalHitRecord> Hits { get; set; } = [];

    public List<string> Exclusions { get; set; } = [];

    public bool EmbeddingsUsed { get; set; }
}

public sealed class PlanningRequestSummary
{
    public string PlanningId { get; set; } = string.Empty;

    public string PlannerVersion { get; set; } = PlanningVersions.PlannerVersion;

    public string CanonicalGoal { get; set; } = string.Empty;

    public string GoalHash { get; set; } = string.Empty;

    public string Provider { get; set; } = string.Empty;

    public string? ModelId { get; set; }

    public string RequestedApprovalMode { get; set; } = string.Empty;

    public string EffectiveApprovalMode { get; set; } = string.Empty;

    public bool ExecuteIfAuthorized { get; set; }

    public bool DryRun { get; set; }

    public DateTimeOffset CreatedAtUtc { get; set; }

    public bool UntrustedContextPresent { get; set; }

    public int UntrustedContextChars { get; set; }
}

public sealed class PlanningValidationReport
{
    public bool SchemaValid { get; set; }

    public bool RefsApproved { get; set; }

    public bool RiskPolicyPassed { get; set; }

    public bool PlanCompiled { get; set; }

    public bool PlanValidatorPassed { get; set; }

    public bool CatalogPreflightPassed { get; set; }

    public string EffectiveApprovalMode { get; set; } = string.Empty;

    public string Risk { get; set; } = string.Empty;

    public List<string> Errors { get; set; } = [];

    public List<string> Warnings { get; set; } = [];

    public string? ProposalHash { get; set; }

    public string? CandidatePlanHash { get; set; }

    public string? RetrievalManifestHash { get; set; }
}

public sealed class PlanningResult
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public string PlanningId { get; init; } = string.Empty;

    public string? ArtifactDirectory { get; init; }

    public PlanningApprovalMode EffectiveApprovalMode { get; init; }

    public PlanProposalDocument? Proposal { get; init; }

    public PlanManifest? CandidatePlan { get; init; }

    public RetrievalManifest? Retrieval { get; init; }

    public PlanningValidationReport Validation { get; init; } = new();

    public PlanningRequestSummary? RequestSummary { get; init; }

    public string? Error { get; init; }

    public bool Executed { get; init; }

    public string? RunId { get; init; }
}

public sealed class PlanningModelRequest
{
    public string CanonicalGoal { get; init; } = string.Empty;

    public string SystemPolicy { get; init; } = string.Empty;

    public string RetrievalContextJson { get; init; } = string.Empty;

    public string? UntrustedContext { get; init; }

    public int MaxOutputBytes { get; init; }

    public int MaxCompletionTokens { get; init; }

    public CancellationToken CancellationToken { get; init; }
}

public sealed class PlanningModelResponse
{
    public string RawJson { get; init; } = string.Empty;

    public string Provider { get; init; } = string.Empty;

    public string? ModelId { get; init; }

    public bool Truncated { get; init; }
}
