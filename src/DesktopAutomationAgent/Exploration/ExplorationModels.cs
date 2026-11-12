using System.Text.Json.Serialization;
using DesktopAutomationAgent.Observation;

namespace DesktopAutomationAgent.Exploration;

public enum ExplorationMode
{
    Observe,
    Navigate
}

public enum ExplorationOutcome
{
    Success,
    NoOp,
    Failure,
    ModalDialog,
    NewWindow,
    ClosedWindow,
    ScopeViolation,
    BudgetExhausted,
    Cancelled,
    StabilizationTimeout,
    StaleRef,
    DryRun
}

public enum ExplorationActionKind
{
    Expand,
    Collapse,
    TabSelect,
    MenuOpen,
    Scroll,
    Focus,
    Hover,
    SelectItem,
    Invoke
}

public sealed class ExplorationPolicyDocument
{
    public const int SchemaVersion = 1;

    public int PolicySchemaVersion { get; set; } = SchemaVersion;

    public string? PolicyId { get; set; }

    public string Mode { get; set; } = "observe";

    public ExplorationApplicationPolicy Application { get; set; } = new();

    public ExplorationScopePolicy Scope { get; set; } = new();

    public List<string> AllowedActionCategories { get; set; } = [];

    public List<ExplorationElementRule> AllowRules { get; set; } = [];

    public List<ExplorationElementRule> DenyRules { get; set; } = [];

    public ExplorationBudgets Budgets { get; set; } = new();

    public bool AllowCloseChildWindowsForBacktrack { get; set; }

    public bool AllowHover { get; set; }

    public bool AllowCoordinateFallback { get; set; }

    public bool AllowDoubleClick { get; set; }
}

public sealed class ExplorationApplicationPolicy
{
    public string ProcessName { get; set; } = string.Empty;

    public string? ExecutablePath { get; set; }

    public string? ExecutableSha256 { get; set; }

    public List<string> AllowedChildProcessNames { get; set; } = [];
}

public sealed class ExplorationScopePolicy
{
    public int? WindowsSessionId { get; set; }

    public List<string> AllowedWindowTitles { get; set; } = [];

    public List<string> DeniedWindowTitles { get; set; } = [];

    public List<string> AllowedWindowClasses { get; set; } = [];

    public List<string> DeniedWindowClasses { get; set; } = [];
}

public sealed class ExplorationElementRule
{
    public string? AutomationId { get; set; }

    public string? Name { get; set; }

    public string? NameContains { get; set; }

    public string? ControlType { get; set; }

    public string? ClassName { get; set; }

    public string? Action { get; set; }

    public string? ActionCategory { get; set; }
}

public sealed class ExplorationBudgets
{
    public int MaxStates { get; set; } = 50;

    public int MaxActions { get; set; } = 100;

    public int MaxDepth { get; set; } = 5;

    public int TotalDurationSeconds { get; set; } = 120;

    public int PerActionTimeoutSeconds { get; set; } = 10;

    public int StabilizationTimeoutSeconds { get; set; } = 5;

    public int StabilizationQuietPeriodMilliseconds { get; set; } = 400;
}

public sealed class ExplorationActionCandidate
{
    public string ActionKey { get; init; } = string.Empty;

    public ExplorationActionKind Kind { get; init; }

    public string Category { get; init; } = string.Empty;

    public string DriverOperation { get; init; } = string.Empty;

    public string ElementKey { get; init; } = string.Empty;

    public string? AutomationId { get; init; }

    public string? Name { get; init; }

    public string? ControlType { get; init; }

    public string? ClassName { get; init; }

    public string SnapshotId { get; init; } = string.Empty;

    public string NodeRef { get; init; } = string.Empty;

    public string StateFingerprint { get; init; } = string.Empty;

    public int Depth { get; init; }

    public int Priority { get; init; }

    public int Ordinal { get; init; }

    /// <summary>
    /// HWND of the window this candidate was resolved from. Before execution the
    /// coordinator switches/revalidates this HWND so sibling-window frontiers cannot
    /// run against the wrong active surface.
    /// </summary>
    public long? TargetHwnd { get; init; }

    /// <summary>
    /// Expected process id for <see cref="TargetHwnd"/> (fail-closed on mismatch).
    /// </summary>
    public int? TargetProcessId { get; init; }

    public Dictionary<string, string>? DriverArguments { get; init; }
}

public sealed class ExplorationStateRecord
{
    public string StateId { get; init; } = string.Empty;

    public string Fingerprint { get; init; } = string.Empty;

    public string? SnapshotId { get; init; }

    public string? WindowTitle { get; init; }

    public string? WindowClassName { get; init; }

    public bool? IsModal { get; init; }

    public int Depth { get; init; }

    public DateTimeOffset CapturedAtUtc { get; init; }

    public int NodeCount { get; init; }
}

public sealed class ExplorationTransitionRecord
{
    public string TransitionId { get; init; } = string.Empty;

    public string FromStateId { get; init; } = string.Empty;

    public string? ToStateId { get; init; }

    public string FromFingerprint { get; init; } = string.Empty;

    public string? ToFingerprint { get; init; }

    public string ActionKey { get; init; } = string.Empty;

    public string ActionKind { get; init; } = string.Empty;

    public string Category { get; init; } = string.Empty;

    public string DriverOperation { get; init; } = string.Empty;

    public string ElementKey { get; init; } = string.Empty;

    public string? AutomationId { get; init; }

    public string? Name { get; init; }

    public ExplorationOutcome Outcome { get; init; }

    public string? Detail { get; init; }

    public DateTimeOffset StartedAtUtc { get; init; }

    public DateTimeOffset CompletedAtUtc { get; init; }

    public int Depth { get; init; }
}

public sealed class ExplorationManifest
{
    public int ManifestSchemaVersion { get; init; } = 1;

    public string ExplorationId { get; init; } = string.Empty;

    public string? PolicyId { get; init; }

    public string Mode { get; init; } = string.Empty;

    public bool DryRun { get; init; }

    public DateTimeOffset StartedAtUtc { get; init; }

    public string? PolicyPath { get; init; }

    public string? PolicySha256 { get; init; }
}

/// <summary>
/// Immutable build identity captured at exploration time. Knowledge merge must use this
/// snapshot rather than re-reading a mutable policy file.
/// </summary>
public sealed class ExplorationBuildIdentity
{
    public const int SchemaVersion = 1;

    public int BuildIdentitySchemaVersion { get; init; } = SchemaVersion;

    public string ProcessName { get; init; } = string.Empty;

    public string? ExecutablePath { get; init; }

    public string? ExecutableSha256 { get; init; }

    public string? ProductVersion { get; init; }

    public string? FileVersion { get; init; }

    public string DriverVersion { get; init; } = string.Empty;

    public int SnapshotSchemaVersion { get; init; } = 1;

    public string ExplorerVersion { get; init; } = "phase4b-1";

    public string PolicySha256 { get; init; } = string.Empty;

    public Dictionary<string, string> ConfigIdentity { get; init; } = new(StringComparer.Ordinal);
}

public sealed class ExplorationSummary
{
    public int SummarySchemaVersion { get; init; } = 1;

    public string ExplorationId { get; init; } = string.Empty;

    public string Status { get; init; } = string.Empty;

    public int ExitCode { get; init; }

    public string Mode { get; init; } = string.Empty;

    public bool DryRun { get; init; }

    public DateTimeOffset StartedAtUtc { get; init; }

    public DateTimeOffset CompletedAtUtc { get; init; }

    public int StatesVisited { get; init; }

    public int ActionsAttempted { get; init; }

    public int TransitionsRecorded { get; init; }

    public int MaxDepthReached { get; init; }

    public string? StopReason { get; init; }

    public List<string> DeniedActionSamples { get; set; } = [];

    public List<string> Warnings { get; set; } = [];

    public string? ArtifactDirectory { get; set; }
}

public sealed record ExplorationRequest
{
    public required string PolicyPath { get; init; }

    public bool DryRun { get; init; }

    public ObservationOptions? ObservationOptions { get; init; }
}

public sealed class ExplorationResult
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public string? Error { get; init; }

    public string? ExplorationId { get; init; }

    public string? ArtifactDirectory { get; init; }

    public ExplorationSummary? Summary { get; init; }

    public List<ExplorationActionCandidate> DryRunCandidates { get; init; } = [];

    [JsonIgnore]
    public SemanticSnapshot? InitialSnapshot { get; init; }
}
