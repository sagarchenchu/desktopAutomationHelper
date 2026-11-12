using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Knowledge;

public static class KnowledgeConstants
{
    public const int KnowledgeSchemaVersion = 1;
    public const int GraphSchemaVersion = 1;
    public const int WorkflowSchemaVersion = 1;
    public const int ApprovalSchemaVersion = 1;

    /// <summary>
    /// Explorer artifact contract version consumed by knowledge merge.
    /// Bump when exploration manifest/state/transition semantics change incompatibly.
    /// </summary>
    public const string ExplorerVersion = "phase4b-1";

    public const int MaxExplorationRuns = 500;
    public const int MaxWindows = 500;
    public const int MaxStates = 5000;
    public const int MaxElements = 20000;
    public const int MaxTransitions = 50000;
    public const int MaxWorkflows = 500;
    public const int MaxFileBytesDefault = 10_485_760;
}

public sealed class ApplicationFingerprint
{
    public string ApplicationId { get; set; } = string.Empty;

    public string? ExecutablePath { get; set; }

    public string? ProductVersion { get; set; }

    public string? FileVersion { get; set; }

    public string? ExecutableSha256 { get; set; }

    public Dictionary<string, string> ConfigIdentity { get; set; } = new(StringComparer.Ordinal);

    public string? DriverVersion { get; set; }

    public int SnapshotSchemaVersion { get; set; } = 1;

    public string ExplorerVersion { get; set; } = KnowledgeConstants.ExplorerVersion;

    public string ExplorationPolicyHash { get; set; } = string.Empty;

    /// <summary>
    /// True when at least one immutable build id is present:
    /// executableSha256, productVersion, or fileVersion.
    /// </summary>
    public bool HasImmutableBuildId() =>
        !string.IsNullOrWhiteSpace(ExecutableSha256)
        || !string.IsNullOrWhiteSpace(ProductVersion)
        || !string.IsNullOrWhiteSpace(FileVersion);
}

public sealed class SourceExplorationRun
{
    public string ExplorationId { get; set; } = string.Empty;

    public string ArtifactHash { get; set; } = string.Empty;

    public string PolicySha256 { get; set; } = string.Empty;

    public string? ManifestPath { get; set; }
}

public sealed class KnowledgeCoverageReport
{
    public int Windows { get; set; }

    public int UniqueStates { get; set; }

    public int ActionableElements { get; set; }

    public int PermittedTransitionsAttempted { get; set; }

    public int SuccessfulTransitions { get; set; }

    public int NoOpTransitions { get; set; }

    public int FailedTransitions { get; set; }

    public int DeniedSamples { get; set; }

    public int UnresolvedElements { get; set; }

    public int Conflicts { get; set; }

    public List<string> TruncationOrBudgetLimits { get; set; } = [];
}

public sealed class KnowledgeIdCollision
{
    public string Kind { get; set; } = string.Empty;

    public string StableId { get; set; } = string.Empty;

    public string LeftKey { get; set; } = string.Empty;

    public string RightKey { get; set; } = string.Empty;
}

public sealed class KnowledgeApproval
{
    public int ApprovalSchemaVersion { get; set; } = KnowledgeConstants.ApprovalSchemaVersion;

    public string Reviewer { get; set; } = string.Empty;

    public DateTimeOffset ReviewedAtUtc { get; set; }

    public string SourceKnowledgeHash { get; set; } = string.Empty;

    public string SourceKnowledgeId { get; set; } = string.Empty;

    public string Decision { get; set; } = "approve";

    public string? Notes { get; set; }

    /// <summary>Always false — knowledge approval never promotes Page Objects.</summary>
    public bool PromotesPageObjects { get; set; }

    public string? ReviewedInputPath { get; set; }

    public string? ReviewedInputHash { get; set; }
}

public sealed class KnowledgeWindowNode
{
    public string WindowId { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    public string? ClassName { get; set; }

    public bool IsModal { get; set; }

    public string Role { get; set; } = "unknown";

    public DateTimeOffset? FirstObservedAtUtc { get; set; }

    public DateTimeOffset? LastObservedAtUtc { get; set; }

    public int ObservationCount { get; set; } = 1;
}

public sealed class KnowledgeStateNode
{
    public string StateId { get; set; } = string.Empty;

    public string SemanticFingerprint { get; set; } = string.Empty;

    public string WindowId { get; set; } = string.Empty;

    public string? Label { get; set; }

    public bool IsEntry { get; set; }

    public bool IsTerminal { get; set; }

    public Dictionary<string, string> ObservedValues { get; set; } = new(StringComparer.Ordinal);

    public List<string> ElementIds { get; set; } = [];

    public double Confidence { get; set; } = 1.0;

    public DateTimeOffset? FirstObservedAtUtc { get; set; }

    public DateTimeOffset? LastObservedAtUtc { get; set; }

    public int ObservationCount { get; set; } = 1;

    public List<string> Provenance { get; set; } = [];
}

public sealed class KnowledgeElementIdentity
{
    public string? AutomationId { get; set; }

    public string? Name { get; set; }

    public string ControlType { get; set; } = string.Empty;

    public string? ClassName { get; set; }

    public string? Role { get; set; }
}

public sealed class KnowledgeLocatorEvidence
{
    public string Kind { get; set; } = "candidate";

    public string Status { get; set; } = "suggested";

    public Dictionary<string, object?> Locator { get; set; } = new(StringComparer.Ordinal);

    public string? ObjectRef { get; set; }
}

public sealed class KnowledgeElementNode
{
    public string ElementId { get; set; } = string.Empty;

    public KnowledgeElementIdentity Identity { get; set; } = new();

    public List<string> Capabilities { get; set; } = [];

    public List<KnowledgeLocatorEvidence> LocatorEvidence { get; set; } = [];

    public string? ApprovedObjectRef { get; set; }

    public double Confidence { get; set; } = 0.5;

    public DateTimeOffset? FirstObservedAtUtc { get; set; }

    public DateTimeOffset? LastObservedAtUtc { get; set; }

    public int ObservationCount { get; set; } = 1;

    public List<string> Provenance { get; set; } = [];
}

public sealed class KnowledgeTransitionEvidence
{
    public string ExplorationId { get; set; } = string.Empty;

    public string? TransitionId { get; set; }

    public string? Detail { get; set; }
}

public sealed class KnowledgeTransitionEdge
{
    public string TransitionId { get; set; } = string.Empty;

    public string Action { get; set; } = string.Empty;

    public string? ActionKind { get; set; }

    public string ActionCategory { get; set; } = string.Empty;

    public Dictionary<string, string> ActionParams { get; set; } = new(StringComparer.Ordinal);

    public string? ElementId { get; set; }

    public string SourceStateId { get; set; } = string.Empty;

    public string? TargetStateId { get; set; }

    public string Outcome { get; set; } = string.Empty;

    public string? ConflictGroupId { get; set; }

    public int ObservationCount { get; set; } = 1;

    public List<KnowledgeTransitionEvidence> Evidence { get; set; } = [];

    public DateTimeOffset? FirstObservedAtUtc { get; set; }

    public DateTimeOffset? LastObservedAtUtc { get; set; }

    public double Confidence { get; set; } = 1.0;
}

public sealed class KnowledgeDialogRelation
{
    public string DialogWindowId { get; set; } = string.Empty;

    public string OwnerWindowId { get; set; } = string.Empty;

    public int ObservationCount { get; set; } = 1;
}

public sealed class KnowledgeChildWindowRelation
{
    public string ChildWindowId { get; set; } = string.Empty;

    public string ParentWindowId { get; set; } = string.Empty;

    public int ObservationCount { get; set; } = 1;
}

public sealed class KnowledgeUnresolvedElement
{
    public string Reason { get; set; } = string.Empty;

    public string RawKey { get; set; } = string.Empty;

    public string? ControlType { get; set; }

    public string? Name { get; set; }
}

public sealed class KnowledgeAmbiguousElement
{
    public string Reason { get; set; } = string.Empty;

    public List<string> CandidateKeys { get; set; } = [];

    public string? StableId { get; set; }
}

public sealed class ApplicationStateGraph
{
    public int GraphSchemaVersion { get; set; } = KnowledgeConstants.GraphSchemaVersion;

    public List<KnowledgeWindowNode> Windows { get; set; } = [];

    public List<KnowledgeStateNode> States { get; set; } = [];

    public List<KnowledgeElementNode> Elements { get; set; } = [];

    public List<KnowledgeTransitionEdge> Transitions { get; set; } = [];

    public List<string> EntryStateIds { get; set; } = [];

    public List<string> TerminalStateIds { get; set; } = [];

    public List<KnowledgeDialogRelation> Dialogs { get; set; } = [];

    public List<KnowledgeChildWindowRelation> ChildWindows { get; set; } = [];

    public List<KnowledgeUnresolvedElement> UnresolvedElements { get; set; } = [];

    public List<KnowledgeAmbiguousElement> AmbiguousElements { get; set; } = [];
}

public sealed class KnowledgeWorkflowStep
{
    public int Ordinal { get; set; }

    public string TransitionId { get; set; } = string.Empty;

    public string Action { get; set; } = string.Empty;

    public string SourceStateId { get; set; } = string.Empty;

    public string? TargetStateId { get; set; }

    public string? Outcome { get; set; }
}

public sealed class ApplicationWorkflow
{
    public int WorkflowSchemaVersion { get; set; } = KnowledgeConstants.WorkflowSchemaVersion;

    public string WorkflowId { get; set; } = string.Empty;

    public string Name { get; set; } = string.Empty;

    public string EntryStateId { get; set; } = string.Empty;

    public string? TerminalStateId { get; set; }

    public List<KnowledgeWorkflowStep> Steps { get; set; } = [];

    public int ObservationCount { get; set; } = 1;

    public double Confidence { get; set; } = 1.0;

    public List<string> Provenance { get; set; } = [];
}

public sealed class ApplicationKnowledgeDocument
{
    public int KnowledgeSchemaVersion { get; set; } = KnowledgeConstants.KnowledgeSchemaVersion;

    public string KnowledgeId { get; set; } = string.Empty;

    public string ApplicationId { get; set; } = string.Empty;

    public string BuildFingerprint { get; set; } = string.Empty;

    public ApplicationFingerprint Fingerprint { get; set; } = new();

    public string Status { get; set; } = "candidate";

    public DateTimeOffset CreatedAtUtc { get; set; }

    public DateTimeOffset? UpdatedAtUtc { get; set; }

    public List<SourceExplorationRun> SourceExplorationRuns { get; set; } = [];

    public ApplicationStateGraph Graph { get; set; } = new();

    public List<ApplicationWorkflow> Workflows { get; set; } = [];

    public KnowledgeCoverageReport Coverage { get; set; } = new();

    public List<KnowledgeIdCollision> IdCollisions { get; set; } = [];

    public KnowledgeApproval? Approval { get; set; }

    public List<string> Warnings { get; set; } = [];
}

public sealed class KnowledgeValidationResult
{
    public bool IsValid => Errors.Count == 0;

    public List<string> Errors { get; } = [];

    public List<string> Warnings { get; } = [];
}

public sealed class KnowledgeCompareResult
{
    public bool CompatibleFingerprints { get; set; }

    public bool GraphsEqual { get; set; }

    public List<string> Differences { get; set; } = [];

    public string? LeftBuildFingerprint { get; set; }

    public string? RightBuildFingerprint { get; set; }
}

public sealed class KnowledgeSearchHit
{
    public string Kind { get; set; } = string.Empty;

    public string Id { get; set; } = string.Empty;

    public string Text { get; set; } = string.Empty;

    public int Score { get; set; }
}

public sealed class KnowledgeBuildResult
{
    public bool Success { get; set; }

    public int ExitCode { get; set; }

    public string? Error { get; set; }

    public string? OutputDirectory { get; set; }

    public ApplicationKnowledgeDocument? Knowledge { get; set; }

    public string? ContentHash { get; set; }

    public List<string> SuggestedPageObjectPaths { get; set; } = [];
}

public sealed class KnowledgeApproveResult
{
    public bool Success { get; set; }

    public int ExitCode { get; set; }

    public string? Error { get; set; }

    public string? OutputDirectory { get; set; }

    public ApplicationKnowledgeDocument? Knowledge { get; set; }
}

public sealed class LoadedExplorationRun
{
    public required string ExplorationId { get; init; }

    public required string RunDirectory { get; init; }

    public required string ArtifactHash { get; init; }

    public required string PolicySha256 { get; init; }

    public required Exploration.ExplorationManifest Manifest { get; init; }

    public required Exploration.ExplorationSummary Summary { get; init; }

    public required IReadOnlyList<Exploration.ExplorationStateRecord> States { get; init; }

    public required IReadOnlyList<Exploration.ExplorationTransitionRecord> Transitions { get; init; }

    public required Exploration.ExplorationBuildIdentity BuildIdentity { get; init; }
}
