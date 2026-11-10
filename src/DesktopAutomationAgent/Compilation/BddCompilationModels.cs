using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Compilation;

public sealed class BddCompileRequest
{
    /// <summary>Path to <c>bdd-action-map.json</c> (workspace-relative or absolute).</summary>
    public required string MapPath { get; init; }

    /// <summary>
    /// Optional recording JSON path. When omitted, <c>sourceRecording</c> from the map
    /// is resolved next to the map file / under the workspace.
    /// </summary>
    public string? RecordingPath { get; init; }

    /// <summary>
    /// Optional output plan path. Defaults to <c>plans/&lt;planId&gt;.plan.json</c>.
    /// </summary>
    public string? OutputPath { get; init; }

    /// <summary>
    /// When true and an object repository path is supplied, emit
    /// <c>$objectRef</c> locators for actions that carry <c>objectRef</c>.
    /// Default is inline locators from the recording (candidates are not auto-promoted).
    /// </summary>
    public bool PreferObjectRefs { get; init; }

    /// <summary>Optional workspace-relative object repository manifest path.</summary>
    public string? ObjectRepositoryPath { get; init; }
}

public sealed class BddCompileResult
{
    public bool Success { get; init; }

    public string MapPath { get; init; } = string.Empty;

    public string RecordingPath { get; init; } = string.Empty;

    public string? OutputPath { get; init; }

    public string? PlanId { get; init; }

    public string? JiraKey { get; init; }

    public string? RecordingId { get; init; }

    public int StepCount { get; init; }

    public int SkippedActionCount { get; init; }

    public int UnmappedEventCount { get; init; }

    public PlanManifest? Plan { get; init; }

    public IReadOnlyList<string> Errors { get; init; } = [];

    public IReadOnlyList<string> Warnings { get; init; } = [];
}

public interface IDeterministicBddPlanCompiler
{
    BddCompileResult Compile(BddCompileRequest request);
}
