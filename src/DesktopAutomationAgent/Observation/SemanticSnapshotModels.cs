using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAutomationAgent.Observation;

public sealed class SemanticSnapshot
{
    public const int SchemaVersion = 1;

    public int SnapshotSchemaVersion { get; init; } = SchemaVersion;

    public string SnapshotId { get; init; } = string.Empty;

    public DateTimeOffset CapturedAtUtc { get; init; }

    public string DriverVersion { get; init; } = string.Empty;

    public string SnapshotHash { get; init; } = string.Empty;

    public ApplicationIdentity Application { get; init; } = new();

    public WindowsSessionIdentity WindowsSession { get; init; } = new();

    public WindowIdentity? Window { get; init; }

    public List<SemanticNode> Nodes { get; init; } = [];

    public bool Truncated { get; init; }

    public List<string> Warnings { get; init; } = [];

    public SnapshotOrdering Ordering { get; init; } = SnapshotOrdering.Default;
}

public sealed class ApplicationIdentity
{
    public int? ProcessId { get; init; }

    public string? ProcessName { get; init; }
}

public sealed class WindowsSessionIdentity
{
    /// <summary>Windows session id for the observed process (not a secret).</summary>
    public int? SessionId { get; init; }
}

public sealed class WindowIdentity
{
    public string? Ref { get; init; }

    public long? Hwnd { get; init; }

    public string? Title { get; init; }

    public string? ClassName { get; init; }

    public int? ProcessId { get; init; }

    public bool? IsModal { get; init; }

    public BoundingRectangleDto? Bounds { get; init; }
}

public sealed class SemanticNode
{
    public string Ref { get; init; } = string.Empty;

    public string? ParentRef { get; init; }

    public int Ordinal { get; init; }

    public int Depth { get; init; }

    public string ControlType { get; init; } = string.Empty;

    public string Role { get; init; } = string.Empty;

    public string? Name { get; init; }

    public string? AutomationId { get; init; }

    public string? ClassName { get; init; }

    public string? HelpText { get; init; }

    public bool Enabled { get; init; } = true;

    public bool Offscreen { get; init; }

    public bool? Focusable { get; init; }

    public bool? HasKeyboardFocus { get; init; }

    public List<string> Patterns { get; init; } = [];

    public List<string> Capabilities { get; init; } = [];

    public BoundingRectangleDto? Bounds { get; init; }

    /// <summary>
    /// Ephemeral runtime id from the driver. Present only for in-memory resolution;
    /// never written into approved knowledge artifacts.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RuntimeId { get; init; }
}

public sealed class BoundingRectangleDto
{
    public double? X { get; init; }

    public double? Y { get; init; }

    public double? Width { get; init; }

    public double? Height { get; init; }

    public double? Left { get; init; }

    public double? Top { get; init; }

    public double? Right { get; init; }

    public double? Bottom { get; init; }
}

public sealed class SnapshotOrdering
{
    public static SnapshotOrdering Default { get; } = new()
    {
        Rule = "depth-first pre-order via parent links; siblings sorted by ordinal then automationId, name, controlType, nodeIndex"
    };

    public string Rule { get; init; } = string.Empty;
}

public sealed record ObservationOptions
{
    public string View { get; init; } = "control";

    public string Root { get; init; } = "activeWindow";

    public int? MaxDepth { get; init; }

    public int? MaxChildren { get; init; }

    public int? MaxNodes { get; init; }

    public bool? IncludeOffscreen { get; init; }

    public bool IncludeBounds { get; init; }

    public bool IncludeHelpText { get; init; }

    public bool IncludeFocusState { get; init; } = true;

    public bool IncludeRuntimeId { get; init; } = true;

    public bool ActionableOnly { get; init; } = true;

    /// <summary>
    /// Optional workspace-relative or absolute path for snapshot.json.
    /// Must not overwrite an existing file.
    /// </summary>
    public string? OutputPath { get; init; }
}

public sealed record ObservationResult
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public string? Error { get; init; }

    public string? ObservationId { get; init; }

    public string? SnapshotFilePath { get; init; }

    public string? SnapshotHash { get; init; }

    public int NodeCount { get; init; }

    public bool Truncated { get; init; }

    public SemanticSnapshot? Snapshot { get; init; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
