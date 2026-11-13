namespace DesktopAutomationDriver.Services.Debug;

/// <summary>Normalized identity fields captured for Debug Mode / vision revalidation.</summary>
public sealed record DebugWindowIdentity
{
    public string? ProcessName { get; init; }
    public int? ProcessId { get; init; }
    public string? WindowTitle { get; init; }
    public IntPtr WindowHandle { get; init; }
    public Rectangle Bounds { get; init; }
    public double DpiScaleX { get; init; } = 1.0;
    public double DpiScaleY { get; init; } = 1.0;
}

/// <summary>One node in a bounded UIA capture for Debug Mode.</summary>
public sealed class DebugUiaNode
{
    public string? AutomationId { get; init; }
    public string? Name { get; init; }
    public string? ClassName { get; init; }
    public string? ControlType { get; init; }
    public int? ControlTypeId { get; init; }
    public string? Path { get; init; }
    public int Depth { get; init; }
    public Rectangle? Bounds { get; init; }
    public IReadOnlyList<string> AvailablePatterns { get; init; } = [];
    public bool IsOffscreen { get; init; }
    public bool IsEnabled { get; init; } = true;
    public bool IsSemanticAccessible { get; init; } = true;
    public string? InaccessibilityReason { get; init; }
}

/// <summary>Bounded capture around the element under the cursor.</summary>
public sealed class DebugTreeCapture
{
    public required DebugUiaNode Target { get; init; }
    public IReadOnlyList<DebugUiaNode> Ancestry { get; init; } = [];
    public IReadOnlyList<DebugUiaNode> Children { get; init; } = [];
    public DebugWindowIdentity? Window { get; init; }
    public Point Cursor { get; init; }
    public DateTimeOffset CapturedAtUtc { get; init; } = DateTimeOffset.UtcNow;
}

/// <summary>Candidate locator generated using object-repository identity rules.</summary>
public sealed class LocatorCandidate
{
    public required string ElementId { get; init; }
    public required IReadOnlyDictionary<string, object?> Locator { get; init; }
    public required string Grade { get; init; }
    public string Strategy { get; init; } = string.Empty;
    public IReadOnlyList<string> Warnings { get; init; } = [];
    public bool AutoPromoted { get; init; }
}

/// <summary>Result of verifying a candidate resolves uniquely.</summary>
public sealed class CandidateVerificationResult
{
    public bool Unique { get; init; }
    public bool Ambiguous { get; init; }
    public bool NotFound { get; init; }
    public int MatchCount { get; init; }
    public string? Message { get; init; }
}

/// <summary>Debug Mode session snapshot for overlays/APIs.</summary>
public sealed class DebugSessionSnapshot
{
    public bool IsActive { get; init; }
    public string? SessionId { get; init; }
    public DateTimeOffset? StartedAtUtc { get; init; }
    public DebugTreeCapture? LastCapture { get; init; }
    public IReadOnlyList<LocatorCandidate> LastCandidates { get; init; } = [];
    public CandidateVerificationResult? LastVerification { get; init; }
    public string? ArtifactPath { get; init; }
    public string StatusText { get; init; } = string.Empty;
}

/// <summary>Hover is a mouse probe — record structure/state/tooltip/no-op outcomes.</summary>
public sealed class HoverProbeResult
{
    public required Point ScreenPoint { get; init; }
    public DebugUiaNode? Before { get; init; }
    public DebugUiaNode? After { get; init; }
    public string? TooltipText { get; init; }
    public bool StructureChanged { get; init; }
    public bool StateChanged { get; init; }
    public bool NoOp { get; init; }
    public DateTimeOffset ProbedAtUtc { get; init; } = DateTimeOffset.UtcNow;
}
