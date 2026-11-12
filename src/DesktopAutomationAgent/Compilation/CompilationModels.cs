using System.Text.Json.Serialization;
using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.Compilation;

public sealed class CompilationResult
{
    public bool Success { get; init; }

    public int ExitCode { get; init; }

    public string? JiraKey { get; init; }

    public string? Fingerprint { get; init; }

    public string? OutputDirectory { get; init; }

    public string? PlanPath { get; init; }

    public string? CompilationPath { get; init; }

    public string? PlanSha256 { get; init; }

    public bool CacheHit { get; init; }

    public PlanManifest? Plan { get; init; }

    public CompilationMetadata? Metadata { get; init; }

    public List<string> Errors { get; init; } = [];

    public List<string> Warnings { get; init; } = [];

    public List<CompilationExplainEntry>? Explain { get; init; }
}

public sealed class CompilationMetadata
{
    public string SchemaVersion { get; init; } = "1";

    public string JiraKey { get; init; } = string.Empty;

    public string Fingerprint { get; init; } = string.Empty;

    public string CompilerVersion { get; init; } = BddCompilerVersions.CompilerVersion;

    public string ParserVersion { get; init; } = BddCompilerVersions.ParserVersion;

    public string InputSha256 { get; init; } = string.Empty;

    public string JiraRevision { get; init; } = string.Empty;

    public string StepRegistrySha256 { get; init; } = string.Empty;

    public string ObjectRepositorySha256 { get; init; } = string.Empty;

    public string? ApprovedKnowledgeSha256 { get; init; }

    public string? CatalogFingerprint { get; init; }

    public string PlanSha256 { get; init; } = string.Empty;

    public bool CacheHit { get; set; }

    /// <summary>Wall-clock metadata only — never feeds the executable plan bytes.</summary>
    public DateTimeOffset CompiledAtUtc { get; init; }

    public List<CompilationStepTrace> Steps { get; init; } = [];
}

public sealed class CompilationStepTrace
{
    public int Ordinal { get; init; }

    public string Keyword { get; init; } = string.Empty;

    public string Text { get; init; } = string.Empty;

    public string MappingId { get; init; } = string.Empty;

    public Dictionary<string, string> Parameters { get; init; } = new(StringComparer.Ordinal);

    public List<string> PlanStepIds { get; init; } = [];
}

public sealed class CompilationExplainEntry
{
    public int Ordinal { get; init; }

    public string Keyword { get; init; } = string.Empty;

    public string Text { get; init; } = string.Empty;

    public string? MappingId { get; init; }

    public string Outcome { get; init; } = string.Empty;

    public List<string> Details { get; init; } = [];
}

public sealed class CompiledArtifactValidationResult
{
    public bool IsValid { get; init; }

    public string? JiraKey { get; init; }

    public string? Fingerprint { get; init; }

    public string? PlanPath { get; init; }

    public string? PlanSha256 { get; init; }

    public List<string> Errors { get; init; } = [];

    public List<string> Warnings { get; init; } = [];
}
