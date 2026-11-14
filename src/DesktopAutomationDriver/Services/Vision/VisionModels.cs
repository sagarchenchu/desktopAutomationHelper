using DesktopAutomationDriver.Services.Debug;

namespace DesktopAutomationDriver.Services.Vision;

/// <summary>Vision fallback policy — disabled by default.</summary>
public sealed class VisionFallbackOptions
{
    /// <summary>Master switch. Default false.</summary>
    public bool Enabled { get; set; }

    /// <summary>
    /// When true, vision may run only if semantic UIA reports the control inaccessible.
    /// Ignored when <see cref="Enabled"/> is false.
    /// </summary>
    public bool AllowWhenUiaInaccessible { get; set; } = true;

    /// <summary>Explicit policy override to permit vision even when UIA is accessible.</summary>
    public bool AllowWhenPolicyPermits { get; set; }

    /// <summary>Maximum age of a screenshot before coordinate actions are rejected.</summary>
    public int MaxScreenshotAgeMilliseconds { get; set; } = 1500;

    /// <summary>When false (default), screenshots with sensitive markers are not persisted.</summary>
    public bool PersistSensitiveScreenshots { get; set; }

    /// <summary>Redact known sensitive tokens from OCR text before persistence.</summary>
    public bool RedactSensitiveText { get; set; } = true;

    /// <summary>Provider: <c>fake</c> (CI) or future real OCR adapter name.</summary>
    public string Provider { get; set; } = "fake";

    /// <summary>Minimum confidence for a vision candidate to be considered.</summary>
    public double MinimumConfidence { get; set; } = 0.6;
}

public sealed class VisionScreenshot
{
    public required byte[] PngBytes { get; init; }
    public required Rectangle WindowBounds { get; init; }
    public required DebugWindowIdentity Window { get; init; }
    public DateTimeOffset CapturedAtUtc { get; init; } = DateTimeOffset.UtcNow;
    public bool ContainsSensitiveContent { get; init; }
    public string? SensitivityReason { get; init; }
    public bool Persisted { get; init; }
    public string? PersistedPath { get; init; }
    public string Scope { get; init; } = "window";
}

public sealed record VisionCandidateTarget
{
    /// <summary>Opaque server-issued id. Absent on client-forged payloads (authorization rejects those).</summary>
    public string? Id { get; init; }
    public required string Label { get; init; }
    public required Rectangle Bounds { get; init; }
    public required double Confidence { get; init; }
    public string Kind { get; init; } = "ocr";
    public bool AutoPromoted { get; init; }
    public required DebugWindowIdentity Window { get; init; }
    /// <summary>Server clock at issue time. Authorization ignores client-supplied timestamps.</summary>
    public DateTimeOffset ProducedAtUtc { get; init; } = DateTimeOffset.UtcNow;
    public string? ScreenshotFingerprint { get; init; }
}

public sealed class CoordinateActionRequest
{
    public required VisionCandidateTarget Candidate { get; init; }
    /// <summary>Must be a fresh server probe — never trust the HTTP caller for this value.</summary>
    public required DebugWindowIdentity CurrentWindow { get; init; }
    /// <summary>Must be server clock (DateTimeOffset.UtcNow) — never trust the HTTP caller.</summary>
    public required DateTimeOffset NowUtc { get; init; }
    public bool IsDestructive { get; init; }
    public string Action { get; init; } = "click";
}

public sealed class CoordinateActionDecision
{
    public bool Allowed { get; init; }
    public string Reason { get; init; } = string.Empty;
    public Point? ScreenPoint { get; init; }
}

public enum VisionEligibility
{
    DisabledByDefault,
    DeniedUiaAccessible,
    AllowedUiaInaccessible,
    AllowedByPolicy,
    Denied
}
